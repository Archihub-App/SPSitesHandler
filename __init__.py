from app.utils.PluginClass import PluginClass
from flask_jwt_extended import jwt_required, get_jwt_identity
from app.utils import DatabaseHandler
from flask import request
from celery import shared_task
from dotenv import load_dotenv
import os
from app.api.records.models import RecordUpdate
from app.api.users.services import has_role
from app.api.adminApi import services
from bson.objectid import ObjectId
import json
from app.utils import DatabaseHandler
from datetime import datetime, timedelta
from app.utils.FernetAuth import fernetAuthenticate
import requests
import msal
 
load_dotenv()
mongodb = DatabaseHandler.DatabaseHandler()
 
USER_FILES_PATH = os.environ.get('USER_FILES_PATH', '')
WEB_FILES_PATH = os.environ.get('WEB_FILES_PATH', '')
ORIGINAL_FILES_PATH = os.environ.get('ORIGINAL_FILES_PATH', '')
TEMPORAL_FILES_PATH = os.environ.get('TEMPORAL_FILES_PATH', '')
plugin_path = os.path.dirname(os.path.abspath(__file__))
CLIENT_ID = os.environ.get("CLIENT_ID", '')
TENANT_ID = os.environ.get("TENANT_ID", '')
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", '')
domain = os.environ.get("SITE_DOMAIN", '')
 
class ExtendedPluginClass(PluginClass):
    def __init__(self, path, import_name, name, description, version, author, type, settings, actions=None, capabilities=None, **kwargs):
        super().__init__(path, __file__, import_name, name, description, version, author, type, settings, actions=actions, capabilities=capabilities, **kwargs)
        
    def download_file(self, drive_id, item_id, file_name):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
        headers = self.get_headers()
        try:
            response = requests.get(url, headers=headers, stream=True)
            response.raise_for_status()
            with open(file_name, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print(f"Downloaded {file_name} successfully to {os.getcwd()}")
        except requests.exceptions.HTTPError as e:
            print(f"Error downloading {file_name}: {e}")
            print(f"Response: {response.text}")

    def get_file_name_by_id(self, drive_id, item_id):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        headers = self.get_headers()
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json().get('name')
    
    def get_headers(self):
        authority = f"https://login.microsoftonline.com/{TENANT_ID}"
        SCOPE = ["https://graph.microsoft.com/.default"]

        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=authority,
            client_credential=CLIENT_SECRET
        )

        result = app.acquire_token_for_client(scopes=SCOPE)

        if "access_token" not in result:
            print(f"Token acquisition failed: {result}")
            if "error_description" in result:
                print(f"Error details: {result['error_description']}")
            raise Exception(f"Failed to acquire token: {result.get('error_description')}")
        
        access_token = result["access_token"]
        headers = {
            'Authorization': f'Bearer {access_token}'
        }
        return headers
 
    def get_folders_content(self, drive_id, folder_id="root", indent=0):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children"
        headers = self.get_headers()
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        items = response.json().get('value', [])
        resp = []
        
        for item in items:
            if item.get('folder'):
                resp.append({"name": item['name'], "type": "folder", "id": item['id']})
                # resp.extend(get_folders_content(drive_id, item['id'], indent + 1))
            elif item.get('file'):
                resp.append({"name": item['name'], "type": "file", "id": item['id']})
        return resp
 
    def get_drive_id(self, site):
        headers = self.get_headers()
        
        url = f"https://graph.microsoft.com/v1.0/sites/{domain}:/sites/{site}"
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        site_data = response.json()
        site_id = site_data.get('id')
        

        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        drive_list = response.json().get('value', [])

        if not drive_list:
            return {'msg': 'No drives found for the specified site.'}
        
        return drive_list
 
    @shared_task(ignore_result=False, name='sharepointSites.bulkUpdate')
    def bulk_update(site, resource_id=None):
        instance = ExtendedPluginClass('sharepointSites','', **plugin_info, isTask=True)
        drive_list = instance.get_drive_id(site)
        headers = instance.get_headers()
        
        drive_id = None
        for drive in drive_list:
            if drive.get('name') == 'Documents':
                drive_id = drive.get('id')
                break
        
        if not drive_id:
            print(f"No 'Documents' drive found for site {site}")
            return {'msg': 'No "Documents" drive found for the specified site.'}
        
        

        content = instance.get_folders_content(drive_id, folder_id='root' if not resource_id else resource_id)
        total_files = 0
        
        for file in content:
            if file['type'] == 'file':
                total_files += 1
                file_name = file['name']
                file_id = file['id']
                output_file = f"{TEMPORAL_FILES_PATH}/{file_name}"
                print(f"Downloading file {file_name} with ID {file_id} to {output_file}")
                instance.download_file(drive_id, file_id, output_file)
                
        return f"Downloaded {total_files} files from SharePoint site {site}."

    def add_routes(self):
        @self.route('/download_resources', methods=['POST'])
        @jwt_required()
        def download_resources():
            try:
                current_user = get_jwt_identity()
                if not has_role(current_user, 'admin') and not has_role(current_user, 'processing'):
                    return {'msg': 'No tiene permisos suficientes'}, 401
 
                current = self.get_plugin_settings()
                if 'sharepoint_site' in current:
                    if current['sharepoint_site'] == '':
                        return {'msg': 'Configuración de SharePoint incompleta'}, 400

                    site = current['sharepoint_site']
                    
                    body = request.get_json()
                    folder_id = body.get('folders_tree', None)
                    if not site:
                        return {'msg': 'Site parameter is required'}, 400
                    if not folder_id:
                        return {'msg': 'Folder ID parameter is required'}, 400
                    
                    folder_id = folder_id[0]['id'] if isinstance(folder_id, list) and len(folder_id) > 0 else folder_id
                    
                    print(f"Downloading resources from SharePoint site: {site}, folder_id: {folder_id}")
                    print(body)
                    task = self.bulk_update.delay(site, folder_id)
                    self.add_task_to_user(task.id, 'SPSitesHandler.bulkUpdate', current_user, 'msg', {
                        'site': site,
                        'folder_id': folder_id
                    })
                    
                    return {'msg': 'Comando enviado para descargars recursos'}, 201
                else:
                    return {'msg': 'Configuración de SharePoint incompleta'}, 400

            except Exception as e:
                print(f"Error en la descarga del recurso: {str(e)}")
                return {'msg': f'Error: {str(e)}'}, 500
            
        @self.route('/get_folder', methods=['POST'])
        @jwt_required()
        def get_folder():
            try:
                current = self.get_plugin_settings()
                if 'sharepoint_site' in current:
                    if current['sharepoint_site'] == '':
                        return {'msg': 'Configuración de SharePoint incompleta'}, 400

                    site = current['sharepoint_site']
                    body = request.get_json()
                    folder_id = body.get('folder_id', 'root')
                    if not site:
                        return {'msg': 'Site parameter is required'}, 400
                    current_user = get_jwt_identity()
                    if not has_role(current_user, 'admin') and not has_role(current_user, 'processing'):
                        return {'msg': 'No tiene permisos suficientes'}, 401
                    
                    drive_list = self.get_drive_id(site)
            
                    drive_id = None
                    for drive in drive_list:
                        if drive.get('name') == 'Documents':
                            drive_id = drive.get('id')
                            break
                    
                    if not drive_id:
                        return {'msg': 'No "Documents" drive found for the specified site.'}
                    
                    

                    content = self.get_folders_content(drive_id, folder_id)
                    tree = [{'id': item['id'], 'name': item['name'], 'post_type': item['type']} for item in content]
                    
                    for r in tree:
                        r['icon'] = 'carpeta' if r['post_type'] == 'folder' else 'archivo'
                        r['children'] = False if r['post_type'] != 'folder' else True

                    return tree, 200
            
                else:
                    return {'msg': 'Configuración de SharePoint incompleta'}, 400

            except Exception as e:
                print(f"Error retrieving folder: {str(e)}")
                return {'msg': f'Error: {str(e)}'}, 500


    @shared_task(ignore_result=False, name='sharepointSites.bulk')
    def bulk(body, user):
        filters = {
            'post_type': body['post_type']
        }
 
        if 'parent' in body:
            if body['parent']:
                filters = {'$or': [{'parents.id': body['parent'], 'post_type': body['post_type']}, {'_id': ObjectId(body['parent'])}]}
 
        if 'resources' in body:
            if body['resources']:
                if len(body['resources']) > 0:
                    filters = {'_id': {'$in': [ObjectId(resource) for resource in body['resources']]}, **filters}
 
        return 'ok'
 
    def get_settings(self):
        @self.route('/settings/<type>', methods=['GET'])
        @jwt_required()
        def get_settings(type):
            try:
                current_user = get_jwt_identity()
 
                if not has_role(current_user, 'admin') and not has_role(current_user, 'processing'):
                    return {'msg': 'No tiene permisos suficientes'}, 401
 
                from app.api.types.services import get_all as get_all_types
                types = get_all_types()
                from app.api.lists.services import get_all as get_all_lists
                lists = get_all_lists()
                lists = lists[0]
 
                if isinstance(types, list):
                    types = tuple(types)[0]
 
                current = self.get_plugin_settings()
                resp = {**self.settings}
                resp = json.loads(json.dumps(resp))
 
                if type == 'all':
                    return resp
                elif type == 'settings':
                    resp['settings'][0]['default'] = current['sharepoint_site'] if 'sharepoint_site' in current else ''
                    resp['settings'][1]['default'] = current['sharepoint_resource_id'] if 'sharepoint_resource_id' in current else ''
                    return resp['settings']
                else:
                    return resp['settings_' + type]
            except Exception as e:
                return {'msg': str(e)}, 500
 
        @self.route('/settings', methods=['POST'])
        @jwt_required()
        def set_settings_update():
            try:
                current_user = get_jwt_identity()
 
                if not has_role(current_user, 'admin') and not has_role(current_user, 'processing'):
                    return {'msg': 'No tiene permisos suficientes'}, 401
 
                body = request.form.to_dict()
                data = body['data']
                data = json.loads(data)
 
                self.set_plugin_settings(data)
 
                return {'msg': 'Configuración guardada'}, 200
 
            except Exception as e:
                return {'msg': str(e)}, 500
 
plugin_info = {
 
    'name': 'Control de SharePoint Sites',
    'description': 'Plugin para acceder a sitios de SharePoint y descargar archivos.',
    'version': '0.1',
    'author': 'BIT SOL SAS',
    'type': ['settings', 'control'],
    'settings': {
        'settings': [
            {
                'type': 'text',
                'label': 'Sitio de SharePoint',
                'id': 'sharepoint_site',
                'default': 'my-site',
                'required': True
            },
            {
                'type': 'text',
                'label': 'ID del recurso a descargar',
                'id': 'sharepoint_resource_id',
                'default': 'my-resource-id',
                'required': False
            }
        ],
        'settings_control': [
            {
                'type':  'instructions',
                'title': 'Instrucciones',
                'text': 'Desde este panel puedes ejecutar tareas de forma manual.'
            },
            {
                'type': 'tree',
                'id': 'folders_tree',
                'api': 'get_folder',
                'label': 'Estructura de carpetas',
                'description': 'Aquí se mostrará la estructura de carpetas del sitio de SharePoint'
            },
            {
                'type': 'button',
                'label': 'Ejecutar',
                'id': 'download_resources',
                'instructions': 'Descargar recursos del sitio de SharePoint de las carpetas seleccionadas.',
            }
        ]
    }
}
 