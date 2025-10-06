# Plugin: Control de SharePoint Sites (SPSitesHandler)

## Descripción
Integra un sitio de SharePoint (Microsoft 365) con la plataforma para:
- Navegar la estructura de carpetas de una biblioteca de documentos.
- Descargar (en lote) archivos y carpetas seleccionadas.
- Crear automáticamente recursos internos (carpetas y documentos) preservando jerarquía.

Usa Microsoft Graph API (flujo client credentials) y ejecuta descargas en segundo plano mediante Celery.

## Características Clave
- Autenticación a Graph usando `msal` (ConfidentialClientApplication).
- Navegación dinámica de carpetas vía endpoint `/get_folder` (modo árbol lazy load).
- Descarga asíncrona disparada desde `/download_resources` -> tarea Celery `sharepointSites.bulkUpdate`.
- Crea recursos tipo `unidad-documental` (archivos) y tipo configurable para carpetas (por defecto `fondo`).
- Evita duplicados: si el archivo ya existe con mismo nombre y hash, omite nueva versión.
- Registro de tarea asociada al usuario que la ejecuta.

## Requisitos
- Python 3.x y dependencias del proyecto (ver `requirements.txt`).
- Celery configurado y en ejecución.
- Variables de entorno de Microsoft 365 (app registrada en Azure AD).
- Roles de usuario: solo usuarios con rol `admin` o `processing` pueden usar los endpoints.

## Variables de Entorno Relevantes
| Variable | Descripción |
|----------|-------------|
| `CLIENT_ID` | ID de la aplicación (Azure AD) |
| `CLIENT_SECRET` | Secreto del cliente (Azure AD) |
| `TENANT_ID` | Tenant ID de Azure AD |
| `SITE_DOMAIN` | Dominio M365 principal (ej: `contoso.sharepoint.com`) |
| `USER_FILES_PATH` | Ruta base de archivos de usuario (no usada directamente aquí) |
| `WEB_FILES_PATH` | Ruta base para archivos web (opcional) |
| `ORIGINAL_FILES_PATH` | Ruta para originales (opcional) |
| `TEMPORAL_FILES_PATH` | Ruta temporal donde se descargan archivos antes de procesar |

## Estructura de Configuración del Plugin
En `plugin_info['settings']` se definen los campos visibles en la UI:
1. `sharepoint_site` (texto) – Nombre del sitio SharePoint (segmento final de la URL del sitio).
2. `sharepoint_drive` (texto) – Nombre de la biblioteca (p.e. `Documentos` / `Documents`).
3. `post_type` (select) – Tipo de contenido para archivos.
4. `folder_post_type` (select) – Tipo de contenido para carpetas.
5. Árbol (`folders_tree`) – Consume `POST /get_folder` para expandir nodos.
6. `sharepoint_resource_id` – ID del recurso raíz donde se crearán los descargados.
7. Botón `download_resources` – Lanza la descarga.

Los selects se llenan dinámicamente desde `types` del sistema.

## Endpoints
Todos requieren JWT + rol permitido.

### POST `/plugins/<plugin_name>/download_resources`
Body JSON esperado:
```json
{
  "folders_tree": [{ "id": "<folderIdSeleccionado>" }],
  "sharepoint_resource_id": "<idRecursoPadre>"
}
```
Acciones:
1. Valida configuración (`sharepoint_site`, `sharepoint_drive`).
2. Lanza tarea Celery `sharepointSites.bulkUpdate` con parámetros (site, folder_id, resource_id, user).
3. Registra la tarea para seguimiento del usuario.

Respuesta 201: `{"msg": "Comando enviado para descargars recursos"}`

### POST `/plugins/<plugin_name>/get_folder`
Body JSON opcional:
```json
{ "folder_id": "<idCarpeta|root>" }
```
Devuelve lista de hijos (carpetas y archivos) con formato árbol minimal:
```json
[{"id": "...", "name": "...", "post_type": "folder|file", "icon": "carpeta|archivo", "children": true|false}]
```

### GET `/plugins/<plugin_name>/settings/<type>`
- `all` devuelve todo el objeto settings.
- `settings` devuelve la sección editable principal.
- `settings_control` devuelve la sección de control.

### POST `/plugins/<plugin_name>/settings`
Formulario (`multipart/form-data`) con campo `data` (JSON serializado) para guardar configuración.

## Tareas Celery
| Nombre | Función | Descripción |
|--------|---------|-------------|
| `sharepointSites.bulkUpdate` | `ExtendedPluginClass.bulk_update` | Recorre recursivamente la carpeta indicada, crea recursos y descarga archivos. |
| `sharepointSites.bulk` | `ExtendedPluginClass.bulk` | Placeholder: actualmente retorna `'ok'` (reservado para futuras operaciones masivas). |

### Flujo Interno de `bulk_update`
1. Obtiene lista de drives del sitio (`/sites/{domain}:/sites/{site}` + `/drives`).
2. Filtra por drive con nombre `'Documentos'` (o el configurado en UI si se amplía en el futuro).
3. Recorre carpeta inicial (o root) usando Graph children endpoint.
4. Para carpetas: crea recurso con `post_type` de carpetas (`folder_post_type` o fijo `fondo` en código actual) y recurre.
5. Para archivos: descarga a `TEMPORAL_FILES_PATH`, calcula hash SHA-256, crea recurso si no existe.
6. El archivo temporal se elimina tras procesar.

## Manejo de Duplicados
Al descargar:
- Busca recurso existente por `metadata.firstLevel.title` + `post_type: 'unidad-documental'`.
- Si existe, obtiene hash (vía `get_hash`) y compara; si coincide, omite subida.
- Si difiere, sube como nueva versión (lógica dependiente de `create_resource`).

## Permisos
Se verifica el usuario actual vía JWT. Debe cumplir:
```
has_role(user, 'admin') OR has_role(user, 'processing')
```
De lo contrario: 401.

## Dependencias Técnicas
- `msal` para Azure AD tokens.
- `requests` para llamadas Graph.
- `celery` para tareas asíncronas.
- `bson` / `ObjectId` para manejo de IDs en Mongo.
- Servicios internos: `resources.create`, `types.get_all`, etc.

## Errores Comunes y Solución
| Situación | Causa Probable | Solución |
|-----------|----------------|----------|
| 400 Configuración incompleta | Campos `sharepoint_site` o `sharepoint_drive` vacíos | Guardar settings antes de usar endpoints |
| 401 No tiene permisos | Usuario sin rol adecuado | Asignar rol `admin` o `processing` |
| No 'Documentos' drive found | Nombre de biblioteca distinto | Ajustar `sharepoint_drive` al nombre real (ej: `Documents`) |
| Token acquisition failed | Credenciales Azure inválidas | Verificar `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID` |
| Recurso no encontrado | `sharepoint_resource_id` inválido | Confirmar ID en la base de datos |

## Buenas Prácticas Operativas
- Usar una cuenta de app con permisos mínimos (Application permissions: Files.Read.All / Sites.Read.All si solo lectura; añadir Files.ReadWrite.All si se requiere escritura futura).
- Mantener limpio `TEMPORAL_FILES_PATH` (el plugin elimina tras uso, pero monitorear espacio).
- Probar primero con una carpeta pequeña antes de ejecutar en raíz.
- Registrar logs de Celery para auditoría de descargas masivas.

## Roadmap Sugerido
- Selección dinámica del tipo `post_type` y `folder_post_type` dentro de la tarea (ahora carpeta fija a `fondo`).
- Soporte para actualización de versiones (metadatos de versión) explícita.
- Paginación de resultados Graph (actualmente asume respuesta única por carpeta).
- Manejo de throttling / retry exponencial.
- Test unitarios para `download_file` y `get_folders_content` con mocks.

## Seguridad
- No almacenar tokens; se solicita uno nuevo por ejecución (`client_credentials`).
- Evitar imprimir secretos en logs.
- Validar longitud / sanitizar nombres de archivos recibidos de SharePoint.

## Ejemplo de Secuencia de Uso
1. Configurar variables de entorno (Azure + rutas).
2. En UI del plugin: completar `sharepoint_site` y `sharepoint_drive` y guardar.
3. Abrir sección de control: expandir árbol (`folders_tree`).
4. Seleccionar carpeta destino y definir `sharepoint_resource_id` (recurso padre ya existente en el sistema).
5. Pulsar botón Ejecutar -> inicia tarea Celery.
6. Monitorear progreso desde la interfaz de tareas (si disponible) o logs.
