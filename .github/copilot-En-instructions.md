## Quick context

This repository is a Google Apps Script (V8) Web App that works as a "Payroll / Attendance System".
- Entry point: the `doGet()` function in `Código.js` which serves `index.html`.
- Main configuration: `appsscript.json` (oauth scopes and `runtimeVersion: V8`).

## Architecture (big-picture)

- Client: HTML+JS in `index.html`, `javascript.html`, `styles.html`. The client calls server functions via `google.script.run` and handles responses with `withSuccessHandler` / `withFailureHandler`.
- Server: Apps Script functions in `Código.js`. Responsibilities: authentication (`verificarAcceso`), sheet management (specific sheets), business logic (attendance, payroll, projects, petty cash), creating Docs/PDFs and sending emails.
- Persistence: a Google Spreadsheet whose ID is read from `PropertiesService` (key `SPREADSHEET_ID`). The expected format/columns are created by `inicializarSistema()` if missing.
- Integrations: Drive (create/share Docs & PDFs), MailApp (send PDFs), PropertiesService, Session (active user email).

## Key files

- `Código.js` — all server-side logic; inspect global constants: `HOJA_COLABORADORES`, `HOJA_ASISTENCIA`, `HOJA_CONFIG`, `HOJA_USUARIOS`, `HOJA_PROYECTOS`, `HOJA_CONTABILIDAD`.
- `index.html` — main template; uses `<?!= include('file') ?>` to inject `javascript.html` and `styles.html`.
- `javascript.html` — client-side UI: `google.script.run` calls, `USER_ROLE` handling and UX flows.
- `appsscript.json` — oauth scopes required (review before adding new APIs).

## Project-specific conventions and patterns

- Implicit sheet schema: many functions assume fixed column positions (e.g. `Colaboradores` expects ID in column A, name in B, email in I). Do not change columns without updating all accesses.
- Date/Currency: dates are formatted as `dd/MM/yyyy` and payroll uses 30 days per month (`valorDia = sueldoBase / 30`).
- Roles: constant `ROLES = { ADMIN: 'ADMINISTRADOR', ASISTENTE: 'ASISTENTE', SIN_ACCESO: 'SIN_ACCESO' }` drives permissions and UI. Many server routes return readable strings for frontend display.
- Messages: server functions often return plain success/error strings (e.g. `"Colaborador registrado correctamente."`) or objects like `{ success: boolean, message, ... }` for complex flows (petty cash).
- Special assignment: the `PROYECTO` option in assignments toggles an additional project selector in the client.

## Client-server flow (concrete examples)

- Register collaborator: the client builds an object and calls `google.script.run.registrarColaborador(colaborador)`; the server validates and does `sheet.appendRow(...)`.
- Save attendance in bulk: the client builds `asistenciasParaGuardar` (items: `colaboradorId, estado, asignacion, vehiculo, horas, observaciones`) and calls `registrarAsistenciasEnLote(asistencias, fecha)`.
- Generate voucher and PDF: `registrarMovimientoCaja(...)` -> `generarValeCaja()` creates a Google Doc, `guardarPdfDesdeDoc()` creates a PDF in Drive which is shared/sent via `MailApp`.

## Important scopes and permissions

Check `appsscript.json`. The included scopes are relevant when testing or adding features:
- spreadsheets, drive, drive.file, documents, script.external_request, gmail.send

If you add integrations, update `oauthScopes` and create a new deployment to trigger the authorization flow.

## Development and deployment (commands & tips)

- The workflow uses `clasp` (you can see `clasp push` usage in this project).
  - Recommended: edit locally, `clasp push` to upload, then deploy from the Apps Script editor or `clasp deploy`.
- To force the authorization prompts (useful when changing scopes or testing Drive/Mail flows), run test functions from the Apps Script editor or call `pruebaGenerarValeMock()` on the server.

## What to check before changing code

1. If you change sheet names, update all references to the `HOJA_*` constants and review `inicializarSistema()` which creates the schema.
2. Do not change expected column formats without migrating data or updating functions that index columns (e.g. `obtenerEmailColaborador` assumes email in column I / index 8).
3. Validations: many functions return readable messages; keep that pattern to avoid breaking the frontend.

## Error handling pattern

- Common pattern: `try { ... } catch (e) { console.error(...); return 'Error...' }`. When you add functions, follow this pattern so the frontend receives consistent messages.

## Quick tips for AI agents

- Read `inicializarSistema()` to learn the minimal spreadsheet structure; use it to spin up a test environment.
- Search for `google.script.run` in `javascript.html` to understand how the client expects responses (plain strings vs objects).
- For data-impacting changes (payroll, vouchers, accounting) add small manual tests and use `pruebaGenerarValeMock()` to validate Drive/Mail permissions.

---
If you want, I can extend this file with sample unit tests or a deployment checklist (clasp + scopes). Which do you prefer?

---

## Deployment checklist and verification

1. Ensure `SPREADSHEET_ID` exists in `PropertiesService` (key: `SPREADSHEET_ID`). If it's missing, run `inicializarSistema()` from the Apps Script editor.
2. Review `appsscript.json` and confirm required scopes. If you add new APIs, update `oauthScopes` and publish a new deployment to prompt authorization.
3. Push changes with `clasp push` and then `clasp deploy` (or deploy from the Apps Script web editor).
4. Run `pruebaGenerarValeMock()` from the editor to force the authorization prompt and verify Drive/Mail flows.
5. Verify the sheets created by `inicializarSistema()` (`Colaboradores`, `RegistrosAsistencia`, `Configuracion`, `Usuarios`, `Project_list`, `contabilidad1`).

## Concrete examples (payloads and calls)

- Register collaborator (client payload):

  {
    id: 'RUT123',
    nombre: 'Juan Perez',
    cargo: 'Operador',
    departamento: 'Produccion',
    fechaIngreso: '2023-06-01',
    sueldoBase: 500000
  }

  Client call: `google.script.run.registrarColaborador(colaborador)`

- Save attendance in bulk (each item):

  [{
    colaboradorId: 'RUT123',
    estado: 'Trabajado',
    asignacion: 'Turno Mañana',
    vehiculo: 'No Aplica',
    horas: 8,
    observaciones: ''
  }, ...]

  Client call: `google.script.run.registrarAsistenciasEnLote(asistencias, fecha)`

- Register petty-cash movement (expense):

  {
    idColaborador: 'RUT123',
    tipoRegistro: 'GASTO',
    tipoMovimiento: 'salida',
    monto: 12345.67,
    detalle: 'Compra materiales'
  }

  Client call: `google.script.run.registrarMovimientoCaja(movimiento)`

## Quick scripts / verifications

1) Basic `SPREADSHEET_ID` verification (Apps Script snippet — run in editor):

```javascript
function verificarSpreadsheetId() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    Logger.log('SPREADSHEET_ID not found in PropertiesService.');
    return false;
  }
  try {
    const ss = SpreadsheetApp.openById(id);
    Logger.log('Spreadsheet accessible: ' + ss.getName());
    return true;
  } catch (e) {
    Logger.log('Could not open the spreadsheet: ' + e.message);
    return false;
  }
}
```

2) Test helper to generate a voucher (force authorization and Drive/Mail flow):

Use the existing `pruebaGenerarValeMock()` function from the Apps Script editor. If you want a verbose variant:

```javascript
function pruebaGenerarValeMockVerbose() {
  const res = pruebaGenerarValeMock();
  Logger.log(JSON.stringify(res));
  return res;
}
```

3) Local verification (PowerShell) — useful commands for developers:

```powershell
# Push local changes to Apps Script
clasp push

# (Optional) Create a new deployment to force scopes
clasp deploy --description "Iter deployment"

# Git push
git add .; git commit -m "Update copilot instructions"; git push
```

4) Quick smoke test (manual):

- Open the Web App URL (if deployed) and confirm the UI loads and access verification appears.
- From the UI: try registering a test collaborator, submit attendance and create a petty-cash movement (small amounts). Confirm rows appear in the corresponding sheets.

---
I added a deployment checklist, examples and verification scripts you can run immediately. Tell me if you'd like unit-test scaffolding or mocks for Drive/Mail flows.

## Sheet schemas (column map)

Note: indexes are 1-based (column A = 1). These expectations are encoded in `inicializarSistema()` and in functions that read/update rows.

- `Colaboradores` (`HOJA_COLABORADORES`)
  1. ID_Colaborador (A) — string
  2. NombreCompleto (B)
  3. Cargo (C)
  4. Departamento (D)
  5. FechaIngreso (E) — date
  6. SueldoBase (F) — number
  7. Estado (G) — 'Activo' / 'Inactivo'
  8. FechaCreacion (H) — timestamp
  9. Email (I) — used by `obtenerEmailColaborador`

- `RegistrosAsistencia` (`HOJA_ASISTENCIA`)
  1. ID_Registro (A) — incremental number
  2. ID_Colaborador (B)
  3. Fecha (C) — date
  4. EstadoAsistencia (D)
  5. Asignacion (E)
  6. Vehiculo (F)
  7. HorasTrabajadas (G)
  8. Observaciones (H)
  9. Timestamp (I)

- `Configuracion` (`HOJA_CONFIG`)
  - Columns used: A..C (Positions: Roles, Departments, AttendanceStates), G (Turn/Assignment/Worksite), K (Vehicle_in_charge), N (Accounting record types). These ranges are used as dropdown sources in the UI.

- `Usuarios` (`HOJA_USUARIOS`)
  1. Email (A)
  2. Rol (B) — expected values: `ADMINISTRADOR`, `ASISTENTE`

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
  (additional columns added by `inicializarSistema()`): `URL_Vale`, `PDF_FileId`, `URL_PDF`, `Vale_Status`

Keep these schemas in mind: changing order or inserting columns without updating `Código.js` will break functions that map fixed indices.
