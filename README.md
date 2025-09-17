# Turnos - Sistema de Nómina (Google Apps Script)

Pequeña guía para configurar, desarrollar y desplegar la Web App basada en Google Apps Script incluida en este repositorio.

## Requisitos

- Node.js + npm (para `clasp`)
- `clasp` instalado globalmente: `npm i -g @google/clasp`
- Acceso a Google account con permisos para editar Apps Script y crear archivos en Drive

## Configuración inicial

1. Clona el repositorio y entra en la carpeta:

```powershell
git clone <repo-url>
cd "C:\Users\juand\OneDrive\Escritorio\GAS-Turnos_SVPZ"
```

2. Inicia sesión con `clasp` (si no lo has hecho):

```powershell
clasp login
```

3. Sube los archivos actuales al proyecto de Apps Script (o verifica con `clasp push` si ya existe vínculo):

```powershell
clasp push
```

4. Asegura la propiedad `SPREADSHEET_ID` en las propiedades del script (desde el editor web de Apps Script o usando script):

- Abre el editor de Apps Script -> Configuración del proyecto -> Propiedades del script -> añade `SPREADSHEET_ID` con el ID de tu spreadsheet.

5. (Opcional) Desde el editor de Apps Script ejecuta `inicializarSistema()` para crear las hojas con el esquema esperado.

## Comandos útiles

- Subir cambios: `clasp push`
- Crear despliegue: `clasp deploy --description "Mi deploy"`
- Forzar autorización (ejecuta desde el editor web): `pruebaGenerarValeMock()` para flujos Drive/Mail

## Verificaciones rápidas

- Verificar `SPREADSHEET_ID` (Apps Script): ver `verificarSpreadsheetId()` en `.github/copilot-instructions.md`.
- Verificar que las hojas existen: `Colaboradores`, `RegistrosAsistencia`, `Configuracion`, `Usuarios`, `Project_list`, `contabilidad1`.

## Esquema de hojas

Los esquemas se describen en `.github/copilot-instructions.md` (sección "Esquemas de hojas (mapa de columnas)"). Manténlos sincronizados si realizas cambios estructurales.

## Desarrollo local y UI

- El frontend está embebido en `index.html` y usa `<?!= include('javascript.html') ?>` para inyectar la lógica y `styles.html` para estilos.
- Para depurar llamadas entre cliente y servidor, usa `Logger.log(...)` en `Código.js` y revisa el registro de ejecuciones en el editor de Apps Script.

## Buenas prácticas específicas del repo

- No cambies índices/columnas de hojas sin actualizar `Código.js`.
- Las funciones deben seguir el patrón `try/catch` y devolver mensajes legibles para el frontend.
- Revisa `appsscript.json` antes de añadir librerías o scopes.

## Soporte

Si quieres que añada scripts de prueba más avanzados o un entorno de pruebas automatizado, indícalo y lo preparo.
