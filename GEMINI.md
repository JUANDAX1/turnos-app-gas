# Gemini Code Assist - Project Guide: Sistema de Nómina

This document provides a comprehensive overview of the "Sistema de Nómina" Google Apps Script project. It is intended to be a guide for AI assistants to understand the project's architecture, data model, and development workflow.

## 1. Project Overview

This is a web application built on Google Apps Script that serves as a Human Resources (HR) and Payroll management system. It allows authorized users to manage employees, track daily attendance, calculate payroll, manage projects, and handle petty cash funds.

**Core Technologies:**
- **Backend:** Google Apps Script (`Código.js`)
- **Database:** Google Sheets
- **Frontend:** HTML, CSS, and JavaScript served via Apps Script's `HtmlService`.
- **Deployment:** Deployed as a Google Apps Script Web App.
- **Local Development:** `clasp` (Command-Line Apps Script Projects) for syncing local files with the Apps Script editor.

## 2. File Structure

- `Código.js`: The main backend file containing all server-side logic written in Google Apps Script (JavaScript). It handles data access, business logic, and serves the frontend.
- `index.html`: The main HTML file for the user interface. It defines the structure of the application, including the sidebar navigation and content tabs.
- `javascript.html`: Contains all client-side JavaScript logic. It is injected into `index.html` and handles UI interactions, form submissions, and communication with the backend via `google.script.run`.
- `styles.html`: Contains all CSS styles for the application. It is injected into `index.html`.
- `appsscript.json`: The manifest file for the Apps Script project. It defines permissions (scopes), dependencies, and other project settings.
- `README.md`: General setup and development guide for human developers.
- `GEMINI.md`: This file. A detailed guide for AI assistants.

## 3. Backend Logic (`Código.js`)

The backend is responsible for all CRUD operations on the Google Sheet, business logic, and user authentication.

### Key Functions & Sections:

- **Configuration (`getSpreadsheetId`, Constants):**
  - `getSpreadsheetId()`: Retrieves the ID of the target Google Sheet from Script Properties, with a hardcoded fallback.
  - `HOJA_*` constants define the names of the sheets used as tables.
  - `ROLES` constant defines user roles (`ADMINISTRADOR`, `ASISTENTE`, `SIN_ACCESO`).

- **Web Server & Authentication:**
  - `doGet()`: The main entry point for the web app. It serves the `index.html` file.
  - `include()`: A utility to inject `javascript.html` and `styles.html` into the main `index.html`.
  - `verificarAcceso()`: Checks the active user's email against the `Usuarios` sheet to determine their role. This is the primary security gate.

- **Module-Specific Logic:**
  - **Gestión de Colaboradores:** `registrarColaborador`, `obtenerColaboradores`.
  - **Lógica de Asistencia:** `obtenerListasParaAsistencia`, `obtenerDatosParaAsistenciaGrid`, `registrarAsistenciasEnLote`, `eliminarRegistroAsistencia`, `actualizarRegistroAsistencia`.
  - **Cálculo de Nómina:** `calcularNomina`.
  - **Gestión de Proyectos:** `registrarProyecto`, `obtenerProyectos`, `actualizarProyecto`, `eliminarProyecto`.
  - **Gestión de Fondos de Caja:** `registrarMovimientoCaja`, `obtenerResumenSaldos`, and helper functions for generating "Vale de Caja" documents (`generarValeCaja`) and sending email notifications.
  - **Bonificaciones:** `obtenerBonificaciones`, `guardarPonderacion`, `obtenerPonderacion`.

- **System Initialization:**
  - `inicializarSistema()`: A crucial function that creates and formats all necessary sheets if they don't exist. It defines the schema for the database.

## 4. Frontend Logic (`javascript.html`)

The frontend is a single-page application (SPA) with a tab-based interface.

- **Initialization:**
  - On `DOMContentLoaded`, `google.script.run.verificarAcceso()` is called.
  - `handleLogin()` processes the response, either showing the app or an "Access Denied" message.
  - `setupUIForRole()` hides/shows navigation items based on the user's role.

- **Communication with Backend:**
  - All calls to the server are made using the `google.script.run` API.
  - Each call has a `.withSuccessHandler()` and often a `.withFailureHandler()` to process the response asynchronously.

- **UI Logic:**
  - `showTab(tabName, ...)`: Manages switching between different content views (Dashboard, Asistencia, etc.).
  - `registrarColaborador(event)`, `solicitarCalculoNomina(event)`, etc.: These functions handle form submissions. They gather data from form inputs, create an object, and send it to the corresponding backend function.
  - `mostrarColaboradores()`, `cargarProyectos()`, etc.: These functions fetch data from the backend and dynamically render it into HTML tables.
  - `showMessage(containerId, ...)`: A utility function to display success, error, or info messages to the user.

## 5. Data Model (Google Sheets Schema)

The `inicializarSistema` function in `Código.js` is the source of truth for the sheet schemas.

- **`Usuarios`**:
  - `A: Email`: User's Google account email.
  - `B: Rol`: User's role (e.g., `ADMINISTRADOR`, `ASISTENTE`).

- **`Colaboradores`**:
  - `A: ID_Colaborador`: Unique ID for the employee.
  - `B: NombreCompleto`: Full name.
  - `C: Cargo`: Job title.
  - `D: Departamento`: Department.
  - `E: FechaIngreso`: Start date.
  - `F: SueldoBase`: Base salary.
  - `G: Estado`: e.g., 'Activo', 'Inactivo'.
  - `H: FechaCreacion`: Timestamp of record creation.
  - `I: Email`: (Inferred from code) Employee's email for notifications.

- **`RegistrosAsistencia`**:
  - `A: ID_Registro`: Unique ID for the attendance record.
  - `B: ID_Colaborador`: Foreign key to `Colaboradores`.
  - `C: Fecha`: Date of the record.
  - `D: EstadoAsistencia`: e.g., 'Trabajado', 'Falta Justificada'.
  - `E: Asignacion`: Task/project/shift assignment.
  - `F: Vehiculo`: Vehicle used, if any.
  - `G: HorasTrabajadas`: Hours worked.
  - `H: Observaciones`: Notes.
  - `I: Timestamp`: Timestamp of record creation.

- **`Configuracion`**:
  - `A: Cargos`: List of job titles.
  - `B: Departamentos`: List of departments.
  - `C: EstadosAsistencia`: List of possible attendance statuses.
  - `G: Turno/Asignacion/Obra`: List of possible assignments.
  - `K: Vehiculo_a_cargo`: List of vehicles.
  - `N: Tipo_registro`: List of transaction types for petty cash (e.g., 'GASTO', 'ANTICIPO').

- **`Project_list`**:
  - `A: project_code`: Unique project code.
  - `B: project_name`: Project name.
  - `C: registration_date`: Project start date.
  - `D: project_address`: Project address.
  - `E: project_georeference`: Geolocation data.
  - `F: project_contact`: Contact person.
  - `G: project_observation`: Notes.
  - `H: Timestamp`: Timestamp of record creation.

- **`contabilidad1`**:
  - `A: ID_Colaborador`: Foreign key to `Colaboradores`.
  - `B: NombreCompleto`: Employee name (denormalized for easy reading).
  - `C: Tipo_registro`: Transaction type from `Configuracion`.
  - `D: ENTRADA_$`: Incoming amount.
  - `E: SALIDA_$`: Outgoing amount.
  - `F: Detalle_transaccion`: Transaction details.
  - `G: Timestamp`: Timestamp of record creation.
  - `H-K`: `URL_Vale`, `PDF_FileId`, `URL_PDF`, `Vale_Status`: Fields related to the auto-generated cash voucher document.

- **`Bonificaciones`**:
  - `A: Proyecto`: Project name.
  - `B...N`: Dynamically generated columns for each collaborator, containing calculated bonus values.

- **`ponderacion`**:
  - `A: Proyecto`: Project name.
  - `B...N`: Columns for each collaborator, containing user-inputted weights/percentages (0-100).

## 6. Development Workflow & Common Tasks

### How to Make a Change
1.  **Local Edit:** Modify the relevant file(s) (`.js`, `.html`) in your local editor.
2.  **Push to Cloud:** Run `clasp push` in your terminal. This uploads the local files to the Apps Script project.
3.  **Test:** Refresh the Web App URL. For backend changes, check the "Executions" log in the Apps Script editor. Use `Logger.log()` for debugging.
4.  **Deploy (for new versions):** Run `clasp deploy` to create a new, immutable version of the app.

### Common Task: Add a New Field to a Form

**Example: Add "Phone Number" to the "Colaboradores" form.**

1.  **`Código.js` (`inicializarSistema`):**
    - Add "Telefono" to the header array for `HOJA_COLABORADORES`.
    - Adjust the range width (`.getRange(1, 1, 1, 9)`).
    - Add formatting if needed.
2.  **`Código.js` (`registrarColaborador`):**
    - Update the `appendRow` call to include `colaborador.telefono || ""`.
3.  **`Código.js` (`obtenerColaboradores`):**
    - Add `telefono: row[X]` to the object being created in the `.map()` function (where X is the new column index).
4.  **`index.html`:**
    - Add a new `<div class="form-group">` for the phone number input inside the `<form id="formColaborador">`.
5.  **`javascript.html` (`registrarColaborador`):**
    - Add `telefono: document.getElementById("telefonoColaborador").value` to the `colaborador` object being created.
6.  **`javascript.html` (`mostrarColaboradores`):**
    - Add a `<th>Teléfono</th>` to the table header.
    - Add a `<td>${col.telefono}</td>` to the table body row creation.
7.  **Push & Test:** Run `clasp push` and test the full flow. You may need to run `inicializarSistema()` manually from the editor once to update the sheet header if it already exists.