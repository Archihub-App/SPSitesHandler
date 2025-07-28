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
 
    @shared_task(ignore_result=False, name='sharepointSites.handler')
    def run_sharepoint_sites_handler(site, resource_id=None):
        import msal
        import requests

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
            print(f"No drives found for site {site}")
            return {'msg': 'No drives found for the specified site.'}
        
        dive_id = None
        for drive in drive_list:
            if drive.get('name') == 'Documents':
                drive_id = drive.get('id')
                break
        
        if not drive_id:
            print(f"No 'Documents' drive found for site {site}")
            return {'msg': 'No "Documents" drive found for the specified site.'}
        
        def print_folder_structure(drive_id, folder_id="root", indent=0):
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children"
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            items = response.json().get('value', [])
            
            for item in items:
                if item.get('folder'):
                    print("  " * indent + f"📁 {item['name']}")
                    print_folder_structure(drive_id, item['id'], indent + 1)
                elif item.get('file'):
                    print("  " * (indent + 1) + f"📄 {item['name']} (ID: {item['id']})")

        print_folder_structure(drive_id)

        def download_file(drive_id, item_id, file_name):
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
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

        def get_file_name_by_id(drive_id, item_id):
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            return response.json().get('name')

        # Example usage: Download a specific file by its ID
        if resource_id:
            file_name = get_file_name_by_id(drive_id, resource_id)
            output_file = f"{TEMPORAL_FILES_PATH}/{file_name}"
            download_file(drive_id, resource_id, output_file)
            print(f"File downloaded to {output_file}")
        else:
            print("No resource ID specified in plugin settings. Please set 'sharepoint_resource_id' to download a specific file.")

    def add_routes(self):
        @self.route('/download_resource', methods=['POST'])
        @jwt_required()
        def download_resource():
            try:
                current_user = get_jwt_identity()
                if not has_role(current_user, 'admin') and not has_role(current_user, 'processing'):
                    return {'msg': 'No tiene permisos suficientes'}, 401
 
                current = self.get_plugin_settings()
                if 'sharepoint_site' in current:
                    if current['sharepoint_site'] == '':
                        return {'msg': 'Configuración de SharePoint incompleta'}, 400

                    site = current['sharepoint_site']

                    self.run_sharepoint_sites_handler.delay(site, current.get('sharepoint_resource_id', None))
                    return {'msg': 'Comando enviado para descargar recurso'}, 201
                else:
                    return {'msg': 'Configuración de SharePoint incompleta'}, 400

            except Exception as e:
                print(f"Error en la descarga del recurso: {str(e)}")
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
                'type': 'button',
                'label': 'Ejecutar',
                'id': 'download_resource',
                'instructions': 'Haz clic en este botón para descargar el recurso especificado en la configuración de SharePoint.',
            }
        ]
    }
}
 