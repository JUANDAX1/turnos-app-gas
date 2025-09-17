## Contexto rápido

Este repositorio es una Web App construida con Google Apps Script (V8) que sirve como "Sistema de Nómina / Asistencia".
- Entrada principal: la función `doGet()` en `Código.js` que sirve `index.html`.
- Configuración principal: `appsscript.json` (scopes y `runtimeVersion: V8`).

## Arquitectura (big-picture)

- Cliente: HTML+JS en `index.html`, `javascript.html`, `styles.html`. El cliente llama al servidor mediante `google.script.run` y maneja respuestas con `withSuccessHandler` / `withFailureHandler`.
- Servidor: funciones de Apps Script en `Código.js`. Maneja: autenticación (`verificarAcceso`), gestión de hojas (hojas concretas), lógica de negocio (asistencia, nómina, proyectos, fondos de caja), creación de Docs/PDFs y envío de correos.
- Persistencia: una Google Spreadsheet cuyo ID se obtiene desde `PropertiesService` (clave `SPREADSHEET_ID`). El formato/columnas esperadas se crean por `inicializarSistema()` si faltan.
- Integraciones: Drive (creación y compartición de Docs/PDF), MailApp (envío de PDFs), PropertiesService, Session (email del usuario).

## Archivos clave

- `Código.js` — toda la lógica del servidor; revisar constantes globales: `HOJA_COLABORADORES`, `HOJA_ASISTENCIA`, `HOJA_CONFIG`, `HOJA_USUARIOS`, `HOJA_PROYECTOS`, `HOJA_CONTABILIDAD`.
- `index.html` — plantilla principal; usa `<?!= include('file') ?>` para inyectar `javascript.html` y `styles.html`.
- `javascript.html` — UI client-side: llamadas a `google.script.run`, manejo de roles (`USER_ROLE`) y flujos de UX.
- `appsscript.json` — scopes requeridos (verlos antes de agregar APIs).

## Convenciones y patrones del proyecto

- Sheet schema implícito: muchas funciones asumen columnas fijas (ej. `Colaboradores` espera ID en columna A, nombre en B, email en I). No cambies columnas sin actualizar los accesos.
- Fecha/Moneda: las fechas se formatean como `dd/MM/yyyy` y la nómina asume 30 días por mes (`valorDia = sueldoBase / 30`).
- Roles: constantes `ROLES = { ADMIN: 'ADMINISTRADOR', ASISTENTE: 'ASISTENTE', SIN_ACCESO: 'SIN_ACCESO' }` controlan permisos y UI. Muchas rutas retornan mensajes legibles para mostrar en el frontend.
- Mensajes: las funciones servidor devuelven frecuentemente strings de éxito/ error (p. ej. `"Colaborador registrado correctamente."`) o `{ success: boolean, message, ... }` para flujos más complejos (fondos de caja).
- Asignación especial: la opción `PROYECTO` en asignaciones activa un selector adicional de proyectos en el cliente.

## Flujo cliente-servidor (ejemplos concretos)

- Registrar colaborador: cliente construye objeto y llama `google.script.run.registrarColaborador(colaborador)`; servidor valida y hace `sheet.appendRow(...)`.
- Guardar asistencias en lote: cliente arma `asistenciasParaGuardar` (cada item: `colaboradorId, estado, asignacion, vehiculo, horas, observaciones`) y llama `registrarAsistenciasEnLote(asistencias, fecha)`.
- Generar vale y PDF: `registrarMovimientoCaja(...)` -> `generarValeCaja()` crea DocumentApp, `guardarPdfDesdeDoc()` crea PDF en Drive y se comparte/envía con `MailApp`.

## Scopes y permisos importantes

Revisar `appsscript.json`. Los scopes incluidos son relevantes cuando se prueba o se agrega funcionalidad:
- spreadsheets, drive, drive.file, documents, script.external_request, gmail.send

Si añades nuevas integraciones revisa y actualiza `appsscript.json` y solicita autorizaciones al desplegar.

## Desarrollo y despliegue (comandos y tips)

- Workflow usado en este repositorio: `clasp` (hay evidencia de `clasp push`).
  - Sugerido: editar localmente, `clasp push` para subir, y desplegar desde el editor de Apps Script o usar `clasp deploy` según sea necesario.
- Para forzar prompts de autorización (útil al cambiar scopes o probar Drive/Mail flows), ejecutar funciones de prueba desde el editor de Apps Script o llamar `pruebaGenerarValeMock()` en el servidor.

## Qué mirar antes de cambiar código

1. Si tocas nombres de hojas, actualiza todas las referencias a las constantes (`HOJA_*`) y revisa `inicializarSistema()` que crea el esquema.
2. No cambies el formato esperado de columnas sin migrar datos o actualizar las funciones que indexan columnas (p. ej. `obtenerEmailColaborador` asume email en columna I / índice 8).
3. Validaciones: muchas funciones retornan mensajes legibles; mantener esa forma evita romper el cliente.

## Errores y manejo esperado

- El patrón común: `try { ... } catch (e) { console.error(...); return 'Error...' }`. Cuando añadas funciones, sigue este patrón para que el frontend reciba mensajes consistentes.

## Sugerencias rápidas para agentes AI

- Lee `inicializarSistema()` para conocer la estructura mínima de la spreadsheet; usa esa función para crear un entorno de pruebas.
- Busca llamadas a `google.script.run` en `javascript.html` para entender cómo espera el cliente las respuestas (strings vs. objetos).
- Para cambios que afectan a datos (nómina, vales, contabilidad) agrega pruebas manuales pequeñas y usa `pruebaGenerarValeMock()` para verificar permisos Drive/Mail.

---
Si quieres, actualizo este archivo con ejemplos de tests/unit (simples) o un checklist de despliegue (clasp + permisos). ¿Qué prefieres que añada? 
---

## Checklist de despliegue y verificación

1. Asegurar `SPREADSHEET_ID` en `PropertiesService` (clave: `SPREADSHEET_ID`). Si falta, ejecutar `inicializarSistema()` desde el editor de Apps Script.
2. Revisar `appsscript.json` y confirmar scopes necesarios. Si añades nuevas APIs, actualiza `oauthScopes` y publica un nuevo despliegue para forzar autorización.
3. Subir cambios con `clasp push` y luego `clasp deploy` (o desplegar desde el editor web de Apps Script).
4. Ejecutar `pruebaGenerarValeMock()` desde el editor para forzar el prompt de autorización y verificar flujo Drive/Mail.
5. Verificar hojas creadas por `inicializarSistema()` (`Colaboradores`, `RegistrosAsistencia`, `Configuracion`, `Usuarios`, `Project_list`, `contabilidad1`).

## Ejemplos concretos (payloads y llamadas)

- Registrar colaborador (payload enviado desde el cliente):

  {
    id: 'RUT123',
    nombre: 'Juan Perez',
    cargo: 'Operador',
    departamento: 'Produccion',
    fechaIngreso: '2023-06-01',
    sueldoBase: 500000
  }

  Llamada cliente: `google.script.run.registrarColaborador(colaborador)`

- Guardar asistencias en lote (cada item):

  [{
    colaboradorId: 'RUT123',
    estado: 'Trabajado',
    asignacion: 'Turno Mañana',
    vehiculo: 'No Aplica',
    horas: 8,
    observaciones: ''
  }, ...]

  Llamada cliente: `google.script.run.registrarAsistenciasEnLote(asistencias, fecha)`

- Registrar movimiento de caja (salida):

  {
    idColaborador: 'RUT123',
    tipoRegistro: 'GASTO',
    tipoMovimiento: 'salida',
    monto: 12345.67,
    detalle: 'Compra materiales' 
  }

  Llamada cliente: `google.script.run.registrarMovimientoCaja(movimiento)`

## Scripts / verificaciones rápidas

1) Verificación básica de `SPREADSHEET_ID` (Apps Script snippet — ejecutar en editor):

```javascript
function verificarSpreadsheetId() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    Logger.log('SPREADSHEET_ID no encontrado en PropertiesService.');
    return false;
  }
  try {
    const ss = SpreadsheetApp.openById(id);
    Logger.log('Spreadsheet accesible: ' + ss.getName());
    return true;
  } catch (e) {
    Logger.log('No se pudo abrir la spreadsheet: ' + e.message);
    return false;
  }
}
```

2) Script de prueba para generar vale (forzar autorización y flujo Drive/Mail):

Usar la función ya existente `pruebaGenerarValeMock()` desde el editor de Apps Script. Si quieres una variante que retorne detalles adicionales:

```javascript
function pruebaGenerarValeMockVerbose() {
  const res = pruebaGenerarValeMock();
  Logger.log(JSON.stringify(res));
  return res;
}
```

3) Verificación local (PowerShell) — comandos útiles para desarrolladores:

```powershell
# Subir cambios locales a Apps Script
clasp push

# (Opcional) Crear un nuevo despliegue para forzar scopes
clasp deploy --description "Iter deployment"

# Verificar git y subir a remoto
git add .; git commit -m "Update copilot instructions"; git push
```

4) Smoke test rápido (manual):

- Ejecutar `doGet()` abriendo la URL del Web App (si ya está desplegada) y verificar que la UI carga y que el mensaje de verificación de acceso aparece.
- Desde la UI: intentar registrar un colaborador de prueba, cargar una asistencia y registrar un movimiento de caja (entrada/salida con monto pequeño). Validar que las filas aparecen en las hojas correspondientes.

---
He añadido checklist, ejemplos y scripts de verificación que puedes ejecutar inmediatamente. Voy a marcar la tarea de feedback como completada una vez confirmes que todo está OK o pidas ajustes.

## Esquemas de hojas (mapa de columnas)

Nota: los índices son 1-based (columna A = 1). Estas expectativas están codificadas en `inicializarSistema()` y en funciones que leen/actualizan filas.

- `Colaboradores` (`HOJA_COLABORADORES`)
  1. ID_Colaborador (A) — string
  2. NombreCompleto (B)
  3. Cargo (C)
  4. Departamento (D)
  5. FechaIngreso (E) — fecha
  6. SueldoBase (F) — número
  7. Estado (G) — 'Activo' / 'Inactivo'
  8. FechaCreacion (H) — timestamp
  9. Email (I) — usado por `obtenerEmailColaborador`

- `RegistrosAsistencia` (`HOJA_ASISTENCIA`)
  1. ID_Registro (A) — número incremental
  2. ID_Colaborador (B)
  3. Fecha (C) — fecha
  4. EstadoAsistencia (D)
  5. Asignacion (E)
  6. Vehiculo (F)
  7. HorasTrabajadas (G)
  8. Observaciones (H)
  9. Timestamp (I)

- `Configuracion` (`HOJA_CONFIG`)
  - Columnas usadas: A..C (Cargos, Departamentos, EstadosAsistencia), G (Turno/Asignacion/Obra), K (Vehiculo_a_cargo), N (Tipos de registro para contabilidad). Contiene listas que se usan como dropdowns en el UI.

- `Usuarios` (`HOJA_USUARIOS`)
  1. Email (A)
  2. Rol (B) — valores esperados: `ADMINISTRADOR`, `ASISTENTE`

- `Project_list` (`HOJA_PROYECTOS`)
  1. project_code (A)
  2. project_name (B)
  3. registration_date (C)
  4. project_address (D)
  5. project_georeference (E)
  6. project_contact (F)
  7. project_observation (G)
  8. Timestamp (H)

- `contabilidad1` (`HOJA_CONTABILIDAD`)
  1. ID_Colaborador
  2. NombreCompleto
  3. Tipo_registro
  4. ENTRADA_$
  5. SALIDA_$
  6. Detalle_transaccion
  7. Timestamp
  (columnas adicionales añadidas por `inicializarSistema()`): `URL_Vale`, `PDF_FileId`, `URL_PDF`, `Vale_Status`

Mantén estos esquemas en mente: cambiar el orden o insertar columnas sin actualizar `Código.js` romperá las funciones que mappean índices fijos.

