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

```markdown
# Instrucciones para agentes (rápido y accionable)

Este repo es una Web App en Google Apps Script (V8) para gestión de nómina/asistencia. Objetivo: dar a un agente AI el contexto mínimo para editar, añadir funciones y desplegar sin romper integraciones.

Arquitectura esencial
- Frontend: `index.html` (plantilla) + `javascript.html` y `styles.html`. El HTML usa `<?!= include('file') ?>` para inyectar partes.
- Backend: `Código.js` contiene todas las funciones publicadas al cliente vía google.script.run.
- Persistencia: una Spreadsheet (ID en `PropertiesService` con clave `SPREADSHEET_ID`). `inicializarSistema()` crea/normaliza hojas y formatos.

Convenciones críticas (no las cambies sin migración)
- Las hojas usan esquemas con índices fijos. Ej.: `Colaboradores` espera ID en A, nombre en B, email en I (índice 8 en 0-based). Revisa `inicializarSistema()` y funciones como `obtenerEmailColaborador`.
- Las funciones devuelven strings legibles o `{ success: boolean, message, ... }`. El frontend asume estas formas para mostrar mensajes.
- Manejo de errores: seguir patrón try/catch -> console.error(...) y retornar mensaje legible.

Llamadas y ejemplos concretos
- Registrar colaborador: cliente -> google.script.run.registrarColaborador(colaborador)
  Ejemplo payload: { id:'RUT123', nombre:'Juan', fechaIngreso:'2023-06-01', sueldoBase:500000 }
- Registrar asistencias en lote: google.script.run.registrarAsistenciasEnLote(asistencias, fecha)
- Registrar movimiento caja (genera Doc + PDF + envía email): google.script.run.registrarMovimientoCaja(movimiento)

Integraciones y permisos
- Revisar `appsscript.json` antes de añadir APIs. Scopes actuales incluyen spreadsheets, drive, documents, drive.file, script.external_request y gmail.send.
- Para forzar autorización de Drive/Mail, ejecutar `pruebaGenerarValeMock()` desde el editor de Apps Script.

Desarrollo y despliegue rápido
- Herramienta: clasp. Flujo típico:
  - Editar localmente
  - clasp push
  - clasp deploy (o desplegar desde el editor web)
- Verificar `SPREADSHEET_ID` en PropertiesService o ejecutar `inicializarSistema()` para crear las hojas.

Dónde buscar al modificar comportamiento
- UI -> `javascript.html` (buscar llamadas `google.script.run` para entender contratos input/output).
- Lógica servidor -> `Código.js` (buscar constantes HOJA_* y funciones: `registrarColaborador`, `registrarAsistenciasEnLote`, `calcularNomina`, `registrarMovimientoCaja`, `inicializarSistema`).

Notas para agentes AI
- No reordenar columnas en hojas sin actualizar todas las lecturas/escrituras en `Código.js`.
- Añade tests manuales mínimos cuando toques generación de Docs/PDFs: usa `pruebaGenerarValeMock()` para validar scope/Drive/Mail.
- Mantén respuestas consistentes (string o objeto con success) para no romper el frontend.

Si algo no está claro, dime qué sección quieres que amplíe (por ejemplo: mapa de columnas completo, checklist de despliegue con comandos, o ejemplos de payloads adicionales).

```
    return false;

  }
