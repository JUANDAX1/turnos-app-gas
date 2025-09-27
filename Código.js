/**
 * @OnlyCurrentDoc
 * Este script se ejecuta automáticamente cuando el documento es abierto.
 * Comprueba si el usuario tiene permiso para ver el contenido.
 * Si no está autorizado, oculta todas las hojas y muestra una de "Acceso Denegado".
 */

// --- CONFIGURACIÓN ---
// ¡IMPORTANTE! Añade aquí los correos de los usuarios que SÍ pueden ver el contenido.
const USUARIOS_AUTORIZADOS = [
  'juandanielcl77@gmail.com',
  'm.melillanca@gmail.com',
  //'otro_usuario_autorizado@email.com'
];

// Nombre de la hoja que se muestra cuando el acceso es denegado.
const HOJA_BLOQUEO = 'Acceso Denegado';


/**
 * Función principal que se ejecuta al abrir el documento.
 * @param {Object} e - Objeto del evento (proporcionado por el activador onOpen).
 */
function onOpen(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const usuarioActual = Session.getActiveUser().getEmail();

  // Comprueba si el usuario actual está en la lista de autorizados.
  if (USUARIOS_AUTORIZADOS.includes(usuarioActual)) {
    // Si está autorizado, nos aseguramos de que pueda ver todas las hojas.
    mostrarHojas(ss);
  } else {
    // Si NO está autorizado, le bloqueamos el acceso.
    ocultarHojas(ss);
    ui.alert(
      'Acceso No Autorizado',
      'No tienes permisos para visualizar el contenido de este documento.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Oculta todas las hojas excepto la de bloqueo.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - La hoja de cálculo activa.
 */
function ocultarHojas(spreadsheet) {
  const todasLasHojas = spreadsheet.getSheets();
  todasLasHojas.forEach(hoja => {
    if (hoja.getName() !== HOJA_BLOQUEO) {
      hoja.hideSheet();
    }
  });
  // Nos aseguramos que la hoja de bloqueo sí esté visible.
  const hojaBloqueo = spreadsheet.getSheetByName(HOJA_BLOQUEO);
  if (hojaBloqueo) {
    hojaBloqueo.showSheet();
  }
}

/**
 * Muestra todas las hojas y oculta la de bloqueo.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - La hoja de cálculo activa.
 */
function mostrarHojas(spreadsheet) {
  const todasLasHojas = spreadsheet.getSheets();
  todasLasHojas.forEach(hoja => {
    if (hoja.getName() === HOJA_BLOQUEO) {
      hoja.hideSheet();
    } else {
      hoja.showSheet();
    }
  });
}



/**
 * @OnlyCurrentDoc
 *
 * El código anterior es una directiva para mejorar el autocompletado de Apps Script.
 */

// ===============================================================
// CONSTANTES GLOBALES Y CONFIGURACIÓN
// ===============================================================

// IMPORTANTE: El ID de la Hoja de Cálculo ahora se gestiona con Propiedades del Script.
// Ve a "Configuración del proyecto" > "Propiedades del script" y añade una propiedad con el nombre "SPREADSHEET_ID" y el valor de tu ID.
function getSpreadsheetId() {
  const SPREADSHEET_ID_HARDCODED = "151hnkLSghwwW54MkgdFa3xj90ICBhv6DakscDcsQvw4"; // Reemplaza con tu ID como respaldo
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID');
    if (spreadsheetId) {
      return spreadsheetId;
    } else {
      // Si no se encuentra la propiedad, se usa el valor hardcodeado y se intenta guardar para el futuro.
      scriptProperties.setProperty('SPREADSHEET_ID', SPREADSHEET_ID_HARDCODED);
      return SPREADSHEET_ID_HARDCODED;
    }
  } catch (e) {
    console.error("No se pudo acceder a las propiedades del script. Usando ID de respaldo.", e);
    return SPREADSHEET_ID_HARDCODED;
  }
}

const HOJA_COLABORADORES = "Colaboradores";
const HOJA_ASISTENCIA = "RegistrosAsistencia";
const HOJA_CONFIG = "Configuracion";
const HOJA_USUARIOS = "Usuarios";
const HOJA_PROYECTOS = "Project_list";
const HOJA_CONTABILIDAD = "contabilidad1";
const HOJA_BONIFICACIONES = "Bonificaciones";
const HOJA_PONDERACION = "ponderacion";
const HOJA_PONDERACION_ESTANDAR = "ponderacion_estandar";

const ROLES = {
  ADMIN: "ADMINISTRADOR",
  ASISTENTE: "ASISTENTE",
  SIN_ACCESO: "SIN_ACCESO"
};

// ===============================================================
// SERVIDOR WEB Y AUTENTICACIÓN
// ===============================================================

function doGet(e) {
  if (e.parameter.page === 'ponderar') {
    return HtmlService.createTemplateFromFile('ponderar.html')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle("Ponderación de Bonos");
  }
  
  // Sirve la página principal por defecto
  return HtmlService.createTemplateFromFile('index.html')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle("Sistema de Gestión de Nómina");
}

function getWebAppUrl(){
  return ScriptApp.getService().getUrl();
}

/**
 * Incluye el contenido de otros archivos en la plantilla HTML.
 * @param {string} filename El nombre del archivo a incluir.
 * @returns {string} El contenido del archivo.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Verifica el email del usuario activo contra la lista de usuarios autorizados.
 * @returns {object} Un objeto con el email y el rol del usuario.
 */
function verificarAcceso() {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetUsuarios = ss.getSheetByName(HOJA_USUARIOS);
    const data = sheetUsuarios.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const userEmail = data[i][0].toString().trim().toLowerCase();
      const userRol = data[i][1].toString().trim().toUpperCase();

      if (userEmail === email.toLowerCase()) {
        return { email: email, rol: userRol };
      }
    }
    
    return { email: email, rol: ROLES.SIN_ACCESO };
  } catch (e) {
    console.error("Error en verificarAcceso:", e);
    return { email: Session.getActiveUser().getEmail(), rol: ROLES.SIN_ACCESO, error: e.message };
  }
}


// ===============================================================
// GESTIÓN DE COLABORADORES
// ===============================================================

/**
 * Registra un nuevo colaborador en la hoja de cálculo.
 * @param {object} colaborador - Objeto con los datos del nuevo colaborador.
 * @returns {string} Un mensaje de éxito o error.
 */
function registrarColaborador(colaborador) {
  try {
    if (!colaborador || !colaborador.id || !colaborador.nombre || !colaborador.sueldoBase) {
      return "Datos incompletos. ID, Nombre y Sueldo Base son obligatorios.";
    }

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_COLABORADORES);

    const ids = sheet.getRange("A2:A").getValues().flat().map(id => id.toString().trim());
    if (ids.includes(colaborador.id.toString().trim())) {
      return "Error: Ya existe un colaborador con ese ID.";
    }
    
    sheet.appendRow([
      colaborador.id.toString().trim(),
      colaborador.nombre,
      colaborador.cargo || "",
      colaborador.departamento || "",
      new Date(colaborador.fechaIngreso),
      parseFloat(colaborador.sueldoBase),
      "Activo",
      new Date()
    ]);
    
    return "Colaborador registrado correctamente.";
  } catch (error) {
    console.error("Error en registrarColaborador:", error);
    return `Error al registrar: ${error.message}`;
  }
}
/**
 * Obtiene la lista completa de colaboradores.
 * @returns {Array<object>} Un arreglo de objetos, donde cada objeto es un colaborador.
 */
function obtenerColaboradores() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_COLABORADORES);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return [];
    }

    const colaboradores = data.slice(1).map(row => ({
      id: row[0],
      nombre: row[1],
      cargo: row[2],
      departamento: row[3],
      fechaIngreso: Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
      sueldoBase: row[5],
      estado: row[6]
    }));
    
    return colaboradores;
  } catch (error) {
    console.error("Error en obtenerColaboradores:", error);
    return [];
  }
}

// ===============================================================
// LÓGICA DE ASISTENCIA Y LISTAS
// ===============================================================

/**
 * Obtiene las listas de colaboradores activos y estados de asistencia para los formularios.
 * @returns {object} Un objeto con arreglos para 'colaboradores' y 'estados'.
 */
function obtenerListasParaAsistencia() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetColaboradores = ss.getSheetByName(HOJA_COLABORADORES);
    const sheetConfig = ss.getSheetByName(HOJA_CONFIG);

    // Obtener colaboradores que están "Activo"
    const colaboradoresData = sheetColaboradores.getDataRange().getValues();
    const colaboradoresActivos = colaboradoresData.slice(1).filter(row => row[6] === 'Activo').map(row => ({
      id: row[0],
      nombre: row[1]
    }));

    // Obtener estados de asistencia desde la configuración
    const estadosData = sheetConfig.getRange("C2:C").getValues();
    const estados = estadosData.flat().filter(String);

    // Obtener asignaciones desde la configuración
    const asignacionesData = sheetConfig.getRange("G2:G").getValues();
    const asignaciones = asignacionesData.flat().filter(String);

    return {
      colaboradores: colaboradoresActivos,
      estados: estados,
      asignaciones: ['PROYECTO', ...asignaciones]
    };
  } catch (error) {
    console.error("Error en obtenerListasParaAsistencia:", error);
    return { colaboradores: [], estados: [], asignaciones: [] };
  }
}

/**
 * Obtiene todos los datos necesarios para construir la parrilla de asistencia.
 * @returns {object} Un objeto con listas de colaboradores, estados y registros del día.
 */
function obtenerDatosParaAsistenciaGrid(fechaSeleccionada) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetColaboradores = ss.getSheetByName(HOJA_COLABORADORES);
    const sheetConfig = ss.getSheetByName(HOJA_CONFIG);
    const sheetAsistencia = ss.getSheetByName(HOJA_ASISTENCIA);
    const sheetProyectos = ss.getSheetByName(HOJA_PROYECTOS);

    // 1. Obtener colaboradores activos
    const colaboradoresData = sheetColaboradores.getDataRange().getValues();
    const colaboradoresActivos = colaboradoresData.slice(1)
      .filter(row => row[6] === 'Activo')
      .map(row => ({ id: row[0], nombre: row[1] }));

    // 2. Obtener listas desde Configuración
    const estados = sheetConfig.getRange("C2:C").getValues().flat().filter(String);
    const asignaciones = sheetConfig.getRange("G2:G").getValues().flat().filter(String);
    const vehiculos = sheetConfig.getRange("K2:K").getValues().flat().filter(String);

    // Obtener lista de proyectos ACTIVOS
    let proyectos = [];
    if (sheetProyectos) {
      // Leemos desde la columna A hasta la G (donde está el estado)
      const proyectosData = sheetProyectos.getRange("A2:G" + sheetProyectos.getLastRow()).getValues(); 
      proyectos = proyectosData
        // Filtramos las filas donde la columna de estado (índice 6) sea "Proyecto Activo"
        .filter(row => row[6] === 'Proyecto Activo') 
        // De las filas filtradas, obtenemos solo el nombre del proyecto (columna B, índice 1)
        .map(row => row[1]); 
    }

    // 3. Obtener registros de la fecha seleccionada
    const fecha = fechaSeleccionada ? new Date(fechaSeleccionada.replace(/-/g, '\/') + ' 00:00:00') : new Date();
    fecha.setHours(0, 0, 0, 0); // Estandarizar a medianoche
    const asistenciaData = sheetAsistencia.getDataRange().getValues();
    const registrosDelDia = asistenciaData.slice(1).filter(row => {
      const fechaRegistro = new Date(row[2]);
      fechaRegistro.setHours(0, 0, 0, 0);
      return fechaRegistro.getTime() === fecha.getTime();
    }).map(row => ({
      idRegistro: row[0],
      colaboradorId: row[1],
      estado: row[3]
    }));

    return {
      colaboradores: colaboradoresActivos,
      estados: estados,
      asignaciones: ['PROYECTO', ...asignaciones],
      vehiculos: vehiculos,
      proyectos: proyectos,
      registrosHoy: registrosDelDia
    };
  } catch (error) {
    console.error("Error en obtenerDatosParaAsistenciaGrid:", error);
    return { error: error.message };
  }
}
// ===============================================================
// LÓGICA DE CÁLCULO DE NÓMINA
// ===============================================================

/**
 * Calcula la pre-nómina para un mes y año específicos.
 * @param {number} mes El mes para el cálculo (1 = Enero, 12 = Diciembre).
 * @param {number} anio El año para el cálculo.
 * @returns {Array<object>} Un arreglo con los resultados de la nómina para cada colaborador.
 */
function calcularNomina(mes, anio) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetColaboradores = ss.getSheetByName(HOJA_COLABORADORES);
    const sheetAsistencia = ss.getSheetByName(HOJA_ASISTENCIA);

    const colaboradoresData = sheetColaboradores.getDataRange().getValues().slice(1);
    const asistenciaData = sheetAsistencia.getDataRange().getValues().slice(1);

    const colaboradoresActivos = colaboradoresData.filter(row => row[6] === 'Activo');
    
    const resultadosNomina = [];

    // Para el cálculo, filtramos los registros de asistencia solo para el mes y año solicitados.
    const mesSeleccionado = parseInt(mes) - 1; // En JavaScript, los meses van de 0 a 11
    const anioSeleccionado = parseInt(anio);

    const asistenciaFiltrada = asistenciaData.filter(row => {
      const fechaRegistro = new Date(row[2]);
      return fechaRegistro.getMonth() === mesSeleccionado && fechaRegistro.getFullYear() === anioSeleccionado;
    });

    for (const colaborador of colaboradoresActivos) {
      const id = colaborador[0];
      const nombre = colaborador[1];
      const sueldoBase = parseFloat(colaborador[5]);

      const registrosDelColaborador = asistenciaFiltrada.filter(row => row[1] == id);

      let diasTrabajados = 0;
      let faltasJustificadas = 0;
      let faltasInjustificadas = 0;
      let licencias = 0;

      for (const registro of registrosDelColaborador) {
        const estado = registro[3].toLowerCase();
        if (estado.includes('trabajado')) {
          diasTrabajados++;
        } else if (estado.includes('justificada')) {
          faltasJustificadas++;
        } else if (estado.includes('injustificada')) {
          faltasInjustificadas++;
        } else if (estado.includes('licencia')) {
          licencias++;
        }
      }

      // Lógica de cálculo del sueldo
      const diasPagables = diasTrabajados + faltasJustificadas + licencias;
      const valorDia = sueldoBase / 30; // Usamos 30 como base de mes comercial
      const sueldoCalculado = valorDia * diasPagables;

      resultadosNomina.push({
        id: id,
        nombre: nombre,
        sueldoBase: sueldoBase,
        diasTrabajados: diasTrabajados,
        faltasJustificadas: faltasJustificadas,
        faltasInjustificadas: faltasInjustificadas,
        licencias: licencias,
        diasPagables: diasPagables,
        sueldoCalculado: sueldoCalculado
      });
    }

    return resultadosNomina;

  } catch (error) {
    console.error("Error en calcularNomina:", error);
    // Devolvemos el mensaje de error para que el frontend pueda mostrarlo
    return { error: `Error en el servidor: ${error.message}` };
  }
}

/**
 * Registra una lista de asistencias en lote.
 * @param {Array<object>} asistencias - Un arreglo de objetos de asistencia.
 * @returns {string} Un mensaje de resumen.
 */
function registrarAsistenciasEnLote(asistencias, fecha) {
  try {
    if (!asistencias || asistencias.length === 0) {
      return "No se enviaron datos para registrar.";
    }
    const fechaRegistro = fecha ? new Date(fecha.replace(/-/g, '\/') + ' 00:00:00') : new Date();

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_ASISTENCIA);
    const registrosActuales = sheet.getDataRange().getValues();
    
    const fechaTexto = Utilities.formatDate(fechaRegistro, Session.getScriptTimeZone(), "dd/MM/yyyy");
    let ultimoId = registrosActuales.length > 0 ? registrosActuales[registrosActuales.length - 1][0] : 0;
    
    const nuevasFilas = [];
    let omitidos = 0;

    for (const asistencia of asistencias) {
      // Verificar duplicados para no registrar dos veces a la misma persona el mismo día
      const yaExiste = registrosActuales.slice(1).some(row => 
        row[1] == asistencia.colaboradorId && 
        Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "dd/MM/yyyy") === fechaTexto
      );

      if (yaExiste) {
        omitidos++;
        continue;
      }
      
      ultimoId++;
      nuevasFilas.push([
        ultimoId,
        asistencia.colaboradorId,
        fechaRegistro,
        asistencia.estado,
        asistencia.asignacion || "",
        asistencia.vehiculo || "",
        asistencia.horas || 8,
        asistencia.observaciones || "",
        new Date()
      ]);
    }

    if (nuevasFilas.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length)
           .setValues(nuevasFilas);
    }
    
    let mensaje = `${nuevasFilas.length} registro(s) guardado(s) correctamente.`;
    if (omitidos > 0) {
      mensaje += ` ${omitidos} se omitieron por ya existir.`;
    }
    return mensaje;

  } catch (error) {
    console.error("Error en registrarAsistenciasEnLote:", error);
    return `Error al guardar: ${error.message}`;
  }
}

/**
 * Obtiene el rol del usuario actual.
 * @returns {string|null} El rol del usuario o null si no se encuentra.
 */
function getRoleForCurrentUser() {
    try {
        const email = Session.getActiveUser().getEmail();
        const ss = SpreadsheetApp.openById(getSpreadsheetId());
        const sheetUsuarios = ss.getSheetByName(HOJA_USUARIOS);
        const data = sheetUsuarios.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            const userEmail = data[i][0].toString().trim().toLowerCase();
            if (userEmail === email.toLowerCase()) {
                return data[i][1].toString().trim().toUpperCase();
            }
        }
        return null;
    } catch (e) {
        console.error("Error en getRoleForCurrentUser:", e);
        return null;
    }
}


/**
 * Elimina un registro de asistencia, aplicando lógica de permisos por rol.
 * @param {object} infoRegistro - Objeto con idRegistro y fechaRegistro.
 * @returns {string} Un mensaje de éxito o error.
 */
function eliminarRegistroAsistencia(infoRegistro) {
  try {
    const rol = getRoleForCurrentUser();
    if (!rol || rol === ROLES.SIN_ACCESO) {
        return "Error: No tiene permisos para realizar esta acción.";
    }

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const fechaRegistro = new Date(infoRegistro.fechaRegistro);
    fechaRegistro.setHours(0, 0, 0, 0);

    if (rol === ROLES.ASISTENTE && fechaRegistro.getTime() !== hoy.getTime()) {
      return "Error: Los asistentes solo pueden editar registros del día actual.";
    }

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_ASISTENCIA);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == infoRegistro.idRegistro) {
        sheet.deleteRow(i + 1);
        return "Registro eliminado. Ahora puede volver a ingresarlo.";
      }
    }
    return "No se encontró el registro para eliminar.";
  } catch (error) {
    console.error("Error en eliminarRegistroAsistencia:", error);
    return `Error al eliminar: ${error.message}`;
  }
}

// ===============================================================
// INICIALIZACIÓN DEL SISTEMA
// ===============================================================

/**
 * Crea y formatea las hojas de cálculo necesarias si no existen.
 */
function inicializarSistema() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());

    // --- Hoja Usuarios ---
    let sheetUsuarios = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheetUsuarios) {
      sheetUsuarios = ss.insertSheet(HOJA_USUARIOS);
      const headers = [["Email", "Rol"]];
      sheetUsuarios.getRange(1, 1, 1, 2).setValues(headers).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
      const currentUserEmail = Session.getActiveUser().getEmail();
      sheetUsuarios.appendRow([currentUserEmail, ROLES.ADMIN]);
      sheetUsuarios.autoResizeColumns(1, 2);
    }
    
    // --- Hoja Colaboradores ---
    let sheetColaboradores = ss.getSheetByName(HOJA_COLABORADORES);
    if (!sheetColaboradores) {
      sheetColaboradores = ss.insertSheet(HOJA_COLABORADORES);
      const headers = [["ID_Colaborador", "NombreCompleto", "Cargo", "Departamento", "FechaIngreso", "SueldoBase", "Estado", "FechaCreacion"]];
      sheetColaboradores.getRange(1, 1, 1, 8).setValues(headers).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
      sheetColaboradores.getRange("A:A").setNumberFormat("@");
      sheetColaboradores.getRange("E:E").setNumberFormat("dd/mm/yyyy");
      sheetColaboradores.getRange("F:F").setNumberFormat("$#,##0");
      sheetColaboradores.getRange("H:H").setNumberFormat("dd/mm/yyyy hh:mm");
      sheetColaboradores.autoResizeColumns(1, 8);
    }

    // --- Hoja RegistrosAsistencia ---
    let sheetAsistencia = ss.getSheetByName(HOJA_ASISTENCIA);
    if (!sheetAsistencia) {
      sheetAsistencia = ss.insertSheet(HOJA_ASISTENCIA);
      const headers = [["ID_Registro", "ID_Colaborador", "Fecha", "EstadoAsistencia", "Asignacion", "Vehiculo", "HorasTrabajadas", "Observaciones", "Timestamp"]];
      sheetAsistencia.getRange(1, 1, 1, 9).setValues(headers).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
      sheetAsistencia.getRange("A:A").setNumberFormat("0");
      sheetAsistencia.getRange("B:B").setNumberFormat("@");
      sheetAsistencia.getRange("C:C").setNumberFormat("dd/mm/yyyy");
      sheetAsistencia.getRange("G:G").setNumberFormat("0.0#");
      sheetAsistencia.getRange("I:I").setNumberFormat("dd/mm/yyyy hh:mm:ss");
      sheetAsistencia.autoResizeColumns(1, 9);
    }

    // --- Hoja Configuracion ---
    let sheetConfig = ss.getSheetByName(HOJA_CONFIG);
    if (!sheetConfig) {
      sheetConfig = ss.insertSheet(HOJA_CONFIG);
      const headers = [["Cargos", "Departamentos", "EstadosAsistencia"]];
      sheetConfig.getRange(1, 1, 1, 3).setValues(headers).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
      sheetConfig.getRange("A2:C5").setValues([
        ["Operador", "Producción", "Trabajado"],
        ["Supervisor", "Calidad", "Falta Justificada"],
        ["Administrativo", "Administración", "Falta Injustificada"],
        ["Gerente", "Gerencia", "Licencia Médica"]
      ]);
      sheetConfig.autoResizeColumns(1, 3);
    }
    
    // Añadir nuevas listas a Configuracion si no existen
    const configHeaders = sheetConfig.getRange(1, 1, 1, sheetConfig.getMaxColumns()).getValues()[0];
    if (configHeaders.indexOf("Turno/Asignacion/Obra") === -1) {
        sheetConfig.getRange(1, 7).setValue("Turno/Asignacion/Obra").setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
        sheetConfig.getRange("G2:G4").setValues([["Turno Mañana"], ["Turno Tarde"], ["Obra Principal"]]);
        sheetConfig.autoResizeColumns(7, 1);
    }
    if (configHeaders.indexOf("Vehiculo_a_cargo") === -1) {
        sheetConfig.getRange(1, 11).setValue("Vehiculo_a_cargo").setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
        sheetConfig.getRange("K2:K4").setValues([["Camioneta 1"], ["Camioneta 2"], ["No Aplica"]]);
        sheetConfig.autoResizeColumns(11, 1);
    }

    // --- Hoja Project_list ---
    let sheetProyectos = ss.getSheetByName(HOJA_PROYECTOS);
    if (!sheetProyectos) {
      sheetProyectos = ss.insertSheet(HOJA_PROYECTOS);
      const headers = [["project_code", "project_name", "registration_date", "client_name", "client_rut", "project_stage", "project_status", "client_phone", "contact_phone", "project_address", "project_georeference", "project_contact", "project_observation", "Timestamp"]];
      sheetProyectos.getRange(1, 1, 1, 14).setValues(headers).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
      sheetProyectos.getRange("A:A").setNumberFormat("@");
      sheetProyectos.getRange("C:C").setNumberFormat("dd/mm/yyyy");
      sheetProyectos.getRange("H:H").setNumberFormat("dd/mm/yyyy hh:mm:ss");
      sheetProyectos.autoResizeColumns(1, 14);
    }

    // --- Hoja Contabilidad ---
    let hojaContabilidad = ss.getSheetByName(HOJA_CONTABILIDAD);
    if (!hojaContabilidad) {
      hojaContabilidad = ss.insertSheet(HOJA_CONTABILIDAD);
      const encabezadosContabilidad = [
        ["ID_Colaborador", "NombreCompleto", "Tipo_registro", "ENTRADA_$", "SALIDA_$", "Detalle_transaccion", "Timestamp"]
      ];
      hojaContabilidad.getRange(1, 1, 1, encabezadosContabilidad[0].length)
        .setValues(encabezadosContabilidad)
        .setBackground("#2E86AB")
        .setFontColor("white")
        .setFontWeight("bold");
      
      // Configurar formatos de columnas
      hojaContabilidad.getRange("A:A").setNumberFormat("@"); // ID como texto
      hojaContabilidad.getRange("D:E").setNumberFormat("#,##0"); // Montos con formato de número
      hojaContabilidad.getRange("G:G").setNumberFormat("dd/mm/yyyy hh:mm:ss"); // Timestamp
      hojaContabilidad.autoResizeColumns(1, 7);
      
      // Configurar validación de datos para Tipo_registro
      const rangoTiposRegistro = ss.getSheetByName(HOJA_CONFIG).getRange("N2:N");
      const regla = SpreadsheetApp.newDataValidation()
        .requireValueInRange(rangoTiposRegistro)
        .setAllowInvalid(false)
        .build();
      hojaContabilidad.getRange("C2:C").setDataValidation(regla);
    }
    
    // Asegurar que existan columnas para almacenar información del vale (documento y PDF) y estado
    const headersCont = hojaContabilidad.getRange(1, 1, 1, hojaContabilidad.getLastColumn()).getValues()[0];
    const neededCols = ['URL_Vale', 'PDF_FileId', 'URL_PDF', 'Vale_Status'];
    neededCols.forEach(colName => {
      if (headersCont.indexOf(colName) === -1) {
        hojaContabilidad.getRange(1, hojaContabilidad.getLastColumn() + 1).setValue(colName)
          .setBackground('#2E86AB').setFontColor('white').setFontWeight('bold');
      }
    });
    hojaContabilidad.autoResizeColumns(1, hojaContabilidad.getLastColumn());
    
    // --- Hoja Bonificaciones ---
    let sheetBonificaciones = ss.getSheetByName(HOJA_BONIFICACIONES);
    if (!sheetBonificaciones) {
      sheetBonificaciones = ss.insertSheet(HOJA_BONIFICACIONES);
      // Cabecera inicial: la primera columna será 'Proyecto'; las columnas de colaboradores se generarán dinámicamente al exportar
      const headers = [["Proyecto"]];
      sheetBonificaciones.getRange(1, 1, 1, 1).setValues(headers).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
      sheetBonificaciones.autoResizeColumns(1, 3);
    }

    // --- Hoja Ponderacion Estandar ---
    let sheetPonderacionEstandar = ss.getSheetByName(HOJA_PONDERACION_ESTANDAR);
    if (!sheetPonderacionEstandar) {
      sheetPonderacionEstandar = ss.insertSheet(HOJA_PONDERACION_ESTANDAR);
      const headers = [["ID_Colaborador", "NombreCompleto", "Cargo", "PonderacionEstandar"]];
      sheetPonderacionEstandar.getRange(1, 1, 1, 4).setValues(headers).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");

      // Poblar con valores iniciales desde la hoja de Colaboradores
      const colaboradoresData = ss.getSheetByName(HOJA_COLABORADORES).getDataRange().getValues().slice(1);
      const datosIniciales = colaboradoresData.map(col => {
        const cargo = (col[2] || '').toLowerCase();
        let ponderacion = 0;
        if (cargo.includes('tecnico') || cargo.includes('técnico')) {
          ponderacion = 65;
        } else if (cargo.includes('ayudante')) {
          ponderacion = 35;
        }
        return [col[0], col[1], col[2], ponderacion];
      });

      if (datosIniciales.length > 0) {
        sheetPonderacionEstandar.getRange(2, 1, datosIniciales.length, 4).setValues(datosIniciales);
      }
      sheetPonderacionEstandar.autoResizeColumns(1, 4);
    }
    
    return "Sistema inicializado correctamente. Todas las hojas han sido creadas y configuradas.";
  } catch (error) {
    console.error("Error en inicializarSistema:", error);
    return `Error al inicializar: ${error.message}`;
  }
}

/**
 * Obtiene la lista de ponderaciones estándar por colaborador.
 * @returns {Array<object>} Un arreglo de objetos con los datos de ponderación estándar.
 */
function obtenerPonderacionEstandar() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_PONDERACION_ESTANDAR);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues().slice(1); // Omitir cabecera
    return data.map(row => ({
      id: row[0],
      nombre: row[1],
      cargo: row[2],
      ponderacion: row[3] || 0
    }));
  } catch (e) {
    console.error("Error en obtenerPonderacionEstandar:", e);
    return { error: e.message };
  }
}

/**
 * Guarda las ponderaciones estándar para todos los colaboradores.
 * @param {Array<object>} ponderaciones - Arreglo de objetos {id, ponderacion}.
 * @returns {object} Mensaje de éxito o error.
 */
function guardarPonderacionEstandar(ponderaciones) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_PONDERACION_ESTANDAR);
    const data = sheet.getRange("A2:A" + sheet.getLastRow()).getValues().flat();
    
    ponderaciones.forEach(item => {
      const rowIndex = data.findIndex(id => id.toString() === item.id.toString());
      if (rowIndex !== -1) {
        // La fila es rowIndex + 2 (porque data empieza en A2 y findIndex es 0-based)
        sheet.getRange(rowIndex + 2, 4).setValue(item.ponderacion);
      }
    });
    
    return { success: true, message: "Ponderaciones estándar guardadas." };
  } catch (e) {
    console.error("Error en guardarPonderacionEstandar:", e);
    return { success: false, message: e.message };
  }
}

/**
 * Consulta los registros de asistencia según un rango de fechas y un colaborador.
 * @param {object} filtros Objeto con fechaDesde, fechaHasta y colaboradorId ('TODOS' para todos).
 * @returns {Array<object>} Un arreglo con los registros encontrados.
 */
function consultarAsistencias(filtros) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetAsistencia = ss.getSheetByName(HOJA_ASISTENCIA);
    const sheetColaboradores = ss.getSheetByName(HOJA_COLABORADORES);

    // Crear un mapa de ID -> Nombre para ser más eficientes
    const colaboradoresData = sheetColaboradores.getDataRange().getValues().slice(1);
    const mapaColaboradores = colaboradoresData.reduce((mapa, fila) => {
      mapa[fila[0]] = fila[1]; // ID -> NombreCompleto
      return mapa;
    }, {});

    const asistenciaData = sheetAsistencia.getDataRange().getValues().slice(1);
    const fechaDesde = new Date(filtros.fechaDesde.replace(/-/g, '\/') + ' 00:00:00');
    const fechaHasta = new Date(filtros.fechaHasta.replace(/-/g, '\/') + ' 23:59:59');

    const resultados = asistenciaData.filter(fila => {
      const fechaRegistro = new Date(fila[2]);
      const idColaborador = fila[1];
      
      const enRangoFecha = fechaRegistro >= fechaDesde && fechaRegistro <= fechaHasta;
      const coincideColaborador = (filtros.colaboradorId === 'TODOS') || (idColaborador == filtros.colaboradorId);

      return enRangoFecha && coincideColaborador;
    }).map(fila => ({
      idRegistro: fila[0],
      idColaborador: fila[1],
      nombreColaborador: mapaColaboradores[fila[1]] || 'Desconocido',
      fecha: Utilities.formatDate(new Date(fila[2]), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      estado: fila[3],
      asignacion: fila[4] || '',
      observaciones: fila[7] || ''
    }));

    return resultados.sort((a, b) => new Date(b.fecha.split('/').reverse().join('-')) - new Date(a.fecha.split('/').reverse().join('-')));

  } catch (e) {
    console.error("Error en consultarAsistencias:", e);
    return { error: e.message };
  }
}

/**
 * Calcula la matriz de bonificaciones (conteo de días asignados por proyecto x colaborador)
 * @param {string} fechaDesde - 'yyyy-mm-dd'
 * @param {string} fechaHasta - 'yyyy-mm-dd'
 * @param {string} filtroBusqueda - texto opcional para filtrar colaboradores (id o nombre parcial)
 * @returns {object} { proyectos: [...], colaboradores: [...], matriz: { proyecto: { colaboradorId: count } } }
 */
function obtenerBonificaciones(fechaDesde, fechaHasta, filtroBusqueda) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetAsistencia = ss.getSheetByName(HOJA_ASISTENCIA);
    const sheetColaboradores = ss.getSheetByName(HOJA_COLABORADORES);

    const colaboradoresData = sheetColaboradores.getDataRange().getValues().slice(1);
    const asistenciaData = sheetAsistencia.getDataRange().getValues().slice(1);

    // Crear mapa ID -> Nombre (trimmed)
    const mapaColaboradores = {};
    colaboradoresData.forEach(r => {
      const id = (r[0] != null) ? r[0].toString().trim() : '';
      const nombre = (r[1] != null) ? r[1].toString().trim() : '';
      if (id) mapaColaboradores[id] = nombre;
    });

    const q = (filtroBusqueda || '').toString().trim().toLowerCase();

    const desde = fechaDesde ? new Date(fechaDesde.replace(/-/g, '/')) : new Date('1970-01-01');
    const hasta = fechaHasta ? new Date(fechaHasta.replace(/-/g, '/')) : new Date();
    hasta.setHours(23,59,59,999);

    const matriz = {};
    const colaboradoresSet = {};

    asistenciaData.forEach(row => {
      try {
        const rawId = row[1] != null ? row[1].toString().trim() : '';
        if (!rawId) return;
        // Aplicar filtro de colaborador si existe
        if (q) {
          const full = (rawId + ' ' + (mapaColaboradores[rawId] || '')).toLowerCase();
          if (full.indexOf(q) === -1) return;
        }

        const fechaRegistro = row[2] ? new Date(row[2]) : null;
        if (!fechaRegistro) return;
        const fr = new Date(fechaRegistro.getFullYear(), fechaRegistro.getMonth(), fechaRegistro.getDate());
        if (fr < new Date(desde.getFullYear(), desde.getMonth(), desde.getDate()) || fr > new Date(hasta.getFullYear(), hasta.getMonth(), hasta.getDate())) return;

        const asignacion = row[4] != null ? row[4].toString().trim() : '';
        if (!asignacion) return;

        // Extraer nombre del proyecto: si contiene 'PROYECTO' intentar obtener la parte posterior, si no usar la asignacion tal cual
        let proyectoNombre = '';
        const up = asignacion.toUpperCase();
        if (up.indexOf('PROYECTO') !== -1) {
          // Buscar ':' y tomar lo que viene después, si no, quitar la palabra PROYECTO y posibles separadores
          const idx = asignacion.indexOf(':');
          if (idx !== -1) proyectoNombre = asignacion.substring(idx + 1).trim();
          else proyectoNombre = asignacion.replace(/PROYECTO\s*-?\s*/i, '').trim();
        } else {
          proyectoNombre = asignacion;
        }

        if (!proyectoNombre) return;

        // Normalizar claves
        const projKey = proyectoNombre;
        const colId = rawId;

        if (!matriz[projKey]) matriz[projKey] = {};
        matriz[projKey][colId] = (matriz[projKey][colId] || 0) + 1;
        colaboradoresSet[colId] = true;
      } catch (inner) {
        // ignorar fila problemática
      }
    });

    // Construir listas ordenadas
    const proyectos = Object.keys(matriz).sort();
    const colaboradores = Object.keys(colaboradoresSet).sort().map(id => ({ id: id, nombre: mapaColaboradores[id] || id }));

    return { proyectos: proyectos, colaboradores: colaboradores, matriz: matriz };
  } catch (e) {
    console.error('Error en obtenerBonificaciones:', e);
    return { error: e.message };
  }
}

/**
 * Escribe la matriz de bonificaciones en la hoja `Bonificaciones`.
 * Filas = proyectos, Columnas = colaboradores (encabezado con 'Nombre (ID)').
 * @param {string} fechaDesde - 'yyyy-mm-dd'
 * @param {string} fechaHasta - 'yyyy-mm-dd'
 * @param {string} filtroBusqueda - texto opcional para filtrar colaboradores
 * @returns {string} Mensaje de resultado
 */
function guardarBonificacionesEnHoja(fechaDesde, fechaHasta, filtroBusqueda) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetBon = ss.getSheetByName(HOJA_BONIFICACIONES) || ss.insertSheet(HOJA_BONIFICACIONES);

    const resultado = obtenerBonificaciones(fechaDesde, fechaHasta, filtroBusqueda);
    if (!resultado || resultado.error) {
      return `Error al calcular bonificaciones: ${resultado && resultado.error ? resultado.error : 'resultado vacío'}`;
    }

    const proyectos = resultado.proyectos || [];
    const colaboradores = resultado.colaboradores || [];
    const matriz = resultado.matriz || {};

    // Limpiar hoja
    sheetBon.clear();

    // Construir cabecera
    const header = ['Proyecto'];
    colaboradores.forEach(c => {
      header.push(`${c.nombre} (${c.id})`);
    });

    const rows = [];
    rows.push(header);

    // Para cada proyecto, construir fila con conteos en orden de colaboradores
    proyectos.forEach(p => {
      const fila = [p];
      colaboradores.forEach(c => {
        const v = (matriz[p] && matriz[p][c.id]) ? matriz[p][c.id] : 0;
        fila.push(v);
      });
      rows.push(fila);
    });

    if (rows.length === 1) {
      // No hay proyectos/colaboradores: dejar una nota
      sheetBon.getRange(1,1).setValue('No hay datos para el periodo o filtros seleccionados.');
      return 'Hoja Bonificaciones actualizada: no hay datos para mostrar.';
    }

    // Escribir matriz en la hoja empezando en A1
    sheetBon.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

    // Formato de cabecera
    sheetBon.getRange(1, 1, 1, rows[0].length).setBackground('#2E86AB').setFontColor('white').setFontWeight('bold');
    sheetBon.autoResizeColumns(1, rows[0].length);

    return `Hoja Bonificaciones actualizada correctamente. Proyectos: ${proyectos.length}, Colaboradores: ${colaboradores.length}`;
  } catch (e) {
    console.error('Error en guardarBonificacionesEnHoja:', e);
    return `Error al guardar bonificaciones: ${e.message}`;
  }
}

/**
 * Actualiza un registro de asistencia existente. Solo para Administradores.
 * @param {object} datos Objeto con idRegistro, nuevoEstado y nuevasObservaciones.
 * @returns {string} Un mensaje de éxito o error.
 */
function actualizarRegistroAsistencia(datos) {
  const rol = getRoleForCurrentUser();
  if (rol !== ROLES.ADMIN) {
    return "Error: Solo los administradores pueden editar registros.";
  }
  
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_ASISTENCIA);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == datos.idRegistro) {
        // Columna D (4) es 'EstadoAsistencia', Columna E (5) es 'Asignacion', Columna H (8) es 'Observaciones'
        sheet.getRange(i + 1, 4).setValue(datos.nuevoEstado);
        sheet.getRange(i + 1, 5).setValue(datos.nuevaAsignacion);
        sheet.getRange(i + 1, 8).setValue(datos.nuevasObservaciones);
        return "Registro actualizado correctamente.";
      }
    }
    return "Error: No se encontró el registro para actualizar.";
  } catch (e) {
    console.error("Error en actualizarRegistroAsistencia:", e);
    return `Error al actualizar: ${e.message}`;
  }
}

// ===============================================================
// GESTIÓN DE PROYECTOS
// ===============================================================

/**
 * Registra un nuevo proyecto en la hoja de cálculo.
 * @param {object} proyecto - Objeto con los datos del nuevo proyecto.
 * @returns {string} Un mensaje de éxito o error.
 */
function registrarProyecto(proyecto) {
  try {
    if (!proyecto || !proyecto.project_code || !proyecto.project_name || !proyecto.registration_date) {
      return "Error: Código, nombre y fecha son campos obligatorios.";
    }

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_PROYECTOS);
    const data = sheet.getDataRange().getValues();
    
    // Verificar si ya existe un proyecto con el mismo código
    const codigoExistente = data.slice(1).some(row => row[0] === proyecto.project_code);
    if (codigoExistente) {
      return "Error: Ya existe un proyecto con ese código.";
    }
    
  sheet.appendRow([
    proyecto.project_code,
    proyecto.project_name,
    new Date(proyecto.registration_date),
    proyecto.client_name || "",
    proyecto.client_rut || "",
    proyecto.project_stage || "",
    proyecto.project_status || "Proyecto Activo",
    proyecto.client_phone || "",
    proyecto.contact_phone || "",
    proyecto.project_address || "",
    proyecto.project_georeference || "",
    proyecto.project_contact || "",
    proyecto.project_observation || "",
    new Date()
  ]);
    
    return "Proyecto registrado correctamente.";
  } catch (error) {
    console.error("Error en registrarProyecto:", error);
    return `Error al registrar: ${error.message}`;
  }
}

/**
 * Obtiene la lista completa de proyectos.
 * @returns {Array<object>} Un arreglo de objetos, donde cada objeto es un proyecto.
 */
function obtenerProyectos() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_PROYECTOS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return [];
    }

    const proyectos = data.slice(1).map(row => ({
      project_code: row[0],
      project_name: row[1],
      registration_date: Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      client_name: row[3] || "",
      client_rut: row[4] || "",
      project_stage: row[5] || "",
      project_status: row[6] || "Proyecto Activo",
      client_phone: row[7] || "",
      contact_phone: row[8] || "",
      project_address: row[9] || "",
      project_georeference: row[10] || "",
      project_contact: row[11] || "",
      project_observation: row[12] || ""
    }));
    
    return proyectos;
  } catch (error) {
    console.error("Error en obtenerProyectos:", error);
    return [];
  }
}

/**
 * Actualiza un proyecto existente.
 * @param {object} proyecto - Objeto con los datos del proyecto a actualizar.
 * @returns {string} Un mensaje de éxito o error.
 */
function actualizarProyecto(proyecto) {
  try {
    if (!proyecto || !proyecto.project_code) {
      return "Error: Se requiere el código del proyecto para actualizar.";
    }

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_PROYECTOS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === proyecto.project_code) {
        sheet.getRange(i + 1, 2).setValue(proyecto.project_name);
        sheet.getRange(i + 1, 3).setValue(new Date(proyecto.registration_date));
        sheet.getRange(i + 1, 4).setValue(proyecto.client_name || "");
        sheet.getRange(i + 1, 5).setValue(proyecto.client_rut || "");
        sheet.getRange(i + 1, 6).setValue(proyecto.project_stage || "");
        sheet.getRange(i + 1, 7).setValue(proyecto.project_status || "Proyecto Activo");
        sheet.getRange(i + 1, 8).setValue(proyecto.client_phone || "");
        sheet.getRange(i + 1, 9).setValue(proyecto.contact_phone || "");
        sheet.getRange(i + 1, 10).setValue(proyecto.project_address || "");
        sheet.getRange(i + 1, 11).setValue(proyecto.project_georeference || "");
        sheet.getRange(i + 1, 12).setValue(proyecto.project_contact || "");
        sheet.getRange(i + 1, 13).setValue(proyecto.project_observation || "");
        
        return "Proyecto actualizado correctamente.";
      }
    }
    
    return "Error: No se encontró el proyecto para actualizar.";
  } catch (error) {
    console.error("Error en actualizarProyecto:", error);
    return `Error al actualizar: ${error.message}`;
  }
}

/**
 * Elimina un proyecto existente.
 * @param {string} projectCode - Código del proyecto a eliminar.
 * @returns {string} Un mensaje de éxito o error.
 */
function eliminarProyecto(projectCode) {
  try {
    if (!projectCode) {
      return "Error: Se requiere el código del proyecto para eliminar.";
    }

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_PROYECTOS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectCode) {
        sheet.deleteRow(i + 1);
        return "Proyecto eliminado correctamente.";
      }
    }
    
    return "Error: No se encontró el proyecto para eliminar.";
  } catch (error) {
    console.error("Error en eliminarProyecto:", error);
    return `Error al eliminar: ${error.message}`;
  }
}

// ===============================================================
// GESTIÓN DE FONDOS DE CAJA
// ===============================================================

/**
 * Obtiene los tipos de registro desde la hoja de configuración.
 * @returns {Array<string>} Lista de tipos de registro.
 */
function obtenerTiposRegistro() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetConfig = ss.getSheetByName(HOJA_CONFIG);
    const tiposRegistro = sheetConfig.getRange("N2:N")
      .getValues()
      .flat()
      .filter(tipo => tipo !== "");
    return tiposRegistro;
  } catch (error) {
    console.error("Error en obtenerTiposRegistro:", error);
    return [];
  }
}

/**
 * Registra un movimiento en la hoja de contabilidad.
 * @param {object} movimiento - Objeto con los datos del movimiento.
 * @returns {string} Mensaje de éxito o error.
 */

function registrarMovimientoCaja(movimiento) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_CONTABILIDAD);
    const sheetColaboradores = ss.getSheetByName(HOJA_COLABORADORES);
    const colaboradoresData = sheetColaboradores.getRange("A:B").getValues();
    const colaborador = colaboradoresData.find(row => row[0].toString() === movimiento.idColaborador.toString());

    if (!colaborador) {
      return { success: false, message: "Error: Colaborador no encontrado." };
    }

    const idRegistroUnico = new Date().getTime().toString(); // ID único para la transacción
    const entrada = movimiento.tipoMovimiento === "entrada" ? movimiento.monto : 0;
    const salida = movimiento.tipoMovimiento === "salida" ? movimiento.monto : 0;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idxIdTransaccion = headers.indexOf('ID_Transaccion');

    if (idxIdTransaccion === -1) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue('ID_Transaccion').setBackground('#2E86AB').setFontColor('white').setFontWeight('bold');
    }

    const nuevaFila = [
      movimiento.idColaborador, colaborador[1], movimiento.tipoRegistro,
      entrada, salida, movimiento.detalle, new Date()
    ];

    const filaInsertada = sheet.getLastRow() + 1;
    sheet.getRange(filaInsertada, 1, 1, nuevaFila.length).setValues([nuevaFila]);

    if (idxIdTransaccion !== -1) {
         sheet.getRange(filaInsertada, idxIdTransaccion + 1).setValue(idRegistroUnico);
    }

    if (movimiento.tipoMovimiento === "salida") {
      const resultVale = generarValeCaja(movimiento, colaborador, idRegistroUnico); // Pasa el ID único
      const carpeta = crearCarpetaValesSiNoExiste();
      const pdfInfo = guardarPdfDesdeDoc(resultVale.fileId, carpeta);

      const idxUrl = headers.indexOf('URL_Vale');
      if (idxUrl !== -1) sheet.getRange(filaInsertada, idxUrl + 1).setValue(resultVale.url);

      const idxPdf = headers.indexOf('PDF_FileId');
      if (idxPdf !== -1) sheet.getRange(filaInsertada, idxPdf + 1).setValue(pdfInfo.fileId);

      const idxUrlPdf = headers.indexOf('URL_PDF');
      if (idxUrlPdf !== -1) sheet.getRange(filaInsertada, idxUrlPdf + 1).setValue(pdfInfo.url);

      return {
        success: true,
        message: 'Movimiento de salida registrado. Oprima los botones para gestionar el vale.',
        idRegistro: idRegistroUnico,
        pdfUrl: pdfInfo.url
      };
    }

    return { success: true, message: "Movimiento de entrada registrado correctamente." };
  } catch (error) {
    console.error("Error en registrarMovimientoCaja:", error);
    return { success: false, message: `Error al registrar el movimiento: ${error.message}` };
  }
}


/**
 * Genera un documento tipo "Vale de Caja" con dos copias (Administrador y Colaborador)
 * y devuelve la URL del documento.
 * @param {object} movimiento
 * @param {Array} colaborador Fila con datos del colaborador [id, nombre]
 * @returns {string} URL del documento creado
 */
function generarValeCaja(movimiento, colaborador) {
  // Crear título y contenido base
  const idCol = movimiento.idColaborador;
  const nombre = colaborador[1] || '';
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const monto = movimiento.monto || 0;
  const detalle = movimiento.detalle || '';

  const titulo = `Vale de Caja - ${idCol} - ${nombre} - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss')}`;

  const doc = DocumentApp.create(titulo);
  const body = doc.getBody();

  // Encabezado general
  const estiloTitulo = {};
  body.appendParagraph('VALE DE CAJA').setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const tabla = body.appendTable();
  let row = tabla.appendTableRow();
  row.appendTableCell('Colaborador');
  row.appendTableCell(nombre);
  row = tabla.appendTableRow();
  row.appendTableCell('ID Colaborador');
  row.appendTableCell(idCol);
  row = tabla.appendTableRow();
  row.appendTableCell('Tipo de Registro');
  row.appendTableCell(movimiento.tipoRegistro || '');
  row = tabla.appendTableRow();
  row.appendTableCell('Fecha');
  row.appendTableCell(fecha);
  row = tabla.appendTableRow();
  row.appendTableCell('Monto entregado');
  row.appendTableCell(`$${Number(monto).toFixed(2)}`);
  row = tabla.appendTableRow();
  row.appendTableCell('Detalle');
  row.appendTableCell(detalle);

  body.appendParagraph('\nDECLARACIÓN:').setBold(true);
  body.appendParagraph('El dinero entregado debe ser rendido o justificado en el plazo establecido por la empresa. Si el dinero no es rendido en el tiempo establecido, éste podrá ser descontado de la remuneración del colaborador según la normativa interna.');

  body.appendParagraph('\n').appendPageBreak();

  // Segunda copia
  body.appendParagraph('COPIA - Colaborador').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const tabla2 = body.appendTable();
  row = tabla2.appendTableRow();
  row.appendTableCell('Colaborador');
  row.appendTableCell(nombre);
  row = tabla2.appendTableRow();
  row.appendTableCell('ID Colaborador');
  row.appendTableCell(idCol);
  row = tabla2.appendTableRow();
  row.appendTableCell('Monto entregado');
  row.appendTableCell(`$${Number(monto).toFixed(2)}`);

  body.appendParagraph('\nFirma del Colaborador: ____________________________');
  body.appendParagraph('\nFirma y Timbre del Administrador: ____________________________');

  doc.saveAndClose();
  const fileId = doc.getId();
  const url = doc.getUrl();
  return { fileId: fileId, url: url };
}

// AGREGA esta nueva función en 'Código.js'

function enviarValePorCorreo(idRegistro) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_CONTABILIDAD);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxIdTransaccion = headers.indexOf('ID_Transaccion');
    const idxPdfId = headers.indexOf('PDF_FileId');
    const idxColaboradorId = headers.indexOf('ID_Colaborador');
    const idxNombre = headers.indexOf('NombreCompleto');
    const idxMontoSalida = headers.indexOf('SALIDA_$');
    const idxDetalle = headers.indexOf('Detalle_transaccion');
    const idxStatus = headers.indexOf('Vale_Status');

    if (idxIdTransaccion === -1 || idxPdfId === -1) {
      return { success: false, message: "Error: Columnas necesarias no encontradas en la hoja." };
    }

    let registroEncontrado = null;
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idxIdTransaccion] == idRegistro) {
        registroEncontrado = data[i];
        rowIndex = i + 1;
        break;
      }
    }

    if (!registroEncontrado) {
      return { success: false, message: "Error: No se encontró el registro del vale." };
    }

    const pdfFileId = registroEncontrado[idxPdfId];
    const idColaborador = registroEncontrado[idxColaboradorId];
    const nombreColaborador = registroEncontrado[idxNombre];
    const monto = registroEncontrado[idxMontoSalida];
    const detalle = registroEncontrado[idxDetalle];

    const emailColaborador = obtenerEmailColaborador(idColaborador);
    const emailAdmin = obtenerEmailAdmin();
    const emailsACompartir = [emailColaborador, emailAdmin].filter(Boolean);

    if (pdfFileId && emailsACompartir.length > 0) {
      compartirArchivoConEmails(pdfFileId, emailsACompartir);

      // --- CORRECCIÓN AQUÍ ---
      const montoFormateado = formatearMonedaCLP(monto); // Usamos la nueva función del servidor
      const subject = `Vale de Caja - ${nombreColaborador} - ${montoFormateado}`;
      const body = `Se ha generado un vale de caja a nombre de ${nombreColaborador} por un monto de ${montoFormateado}.\n\nDetalle: ${detalle}\n\nEl PDF se encuentra adjunto.`;
      // --- FIN DE LA CORRECCIÓN ---

      if (emailColaborador) enviarCorreoConAdjunto(emailColaborador, subject, body, pdfFileId);
      if (emailAdmin) enviarCorreoConAdjunto(emailAdmin, subject, body, pdfFileId);

      if (idxStatus !== -1 && rowIndex !== -1) {
        sheet.getRange(rowIndex, idxStatus + 1).setValue('ENVIADO');
      }
      return { success: true, message: `Correo enviado a: ${emailsACompartir.join(', ')}` };
    }

    return { success: false, message: "No se encontró ID de PDF o destinatarios para enviar el correo." };
  } catch (e) {
    console.error("Error en enviarValePorCorreo:", e);
    return { success: false, message: e.message };
  }
}


/**
 * Obtiene los días trabajados por proyecto y colaborador desde la hoja 'Bonificaciones'.
 * La cabecera de la hoja debe tener el formato 'Nombre (ID)'.
 * @returns {object} Objeto con la estructura { [nombreProyecto]: { [idColaborador]: dias } }
 */
function obtenerDiasTrabajadosBonos() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_BONIFICACIONES);
    if (!sheet) {
      // Si la hoja no existe, podría ser porque aún no se ha corrido el cálculo de bonificaciones.
      // Devolver un objeto vacío es un caso manejable en el frontend.
      return {}; 
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return {}; // No hay datos
    }

    const headers = data[0];
    const mapaColaboradores = {}; // Mapa de indice de columna a ID de colaborador

    // Procesar cabeceras para extraer IDs de colaboradores
    // Se asume formato "Nombre (ID)" a partir de la segunda columna
    for (let i = 1; i < headers.length; i++) {
      const header = headers[i];
      const match = header.match(/\(([^)]+)\)$/); // Extraer contenido del paréntesis al final
      if (match && match[1]) {
        mapaColaboradores[i] = match[1].trim(); // { "1": "C001", "2": "C002", ... }
      }
    }

    const diasTrabajados = {};
    // Procesar filas de datos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nombreProyecto = row[0];
      if (!nombreProyecto) continue; // Omitir filas sin nombre de proyecto

      diasTrabajados[nombreProyecto] = {};
      for (let j = 1; j < row.length; j++) {
        const colabId = mapaColaboradores[j];
        if (colabId) {
          diasTrabajados[nombreProyecto][colabId] = Number(row[j]) || 0;
        }
      }
    }

    return diasTrabajados;
  } catch (e) {
    console.error("Error en obtenerDiasTrabajadosBonos:", e);
    return { error: e.message };
  }
}

/**
 * Obtiene el resumen de saldos por colaborador.
 * @returns {Array<object>} Arreglo con los saldos por colaborador.
 */
/**
 * Obtiene el resumen de saldos por colaborador.
 * Ahora acepta parámetros opcionales para búsqueda inteligente:
 * - filtro: búsqueda parcial en id, nombre y tipoRegistro (case-insensitive)
 * - nombreColaborador: filtro específico por nombre (parcial, case-insensitive)
 * @param {object|string} opciones Puede ser un string (filtro) o un objeto { filtro, nombreColaborador }
 * @returns {Array<object>} Arreglo con los saldos por colaborador filtrados
 */
function obtenerResumenSaldos(opciones) {
  try {
    // Normalizar parámetros
    let filtro = '';
    let nombreColaborador = '';
    if (typeof opciones === 'string') {
      filtro = opciones;
    } else if (typeof opciones === 'object' && opciones !== null) {
      filtro = opciones.filtro || '';
      nombreColaborador = opciones.nombreColaborador || '';
    }

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_CONTABILIDAD);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return [];
    }

    const saldos = {};

    // Procesar cada registro
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = row[0] != null ? row[0].toString() : '';
      const nombre = row[1] != null ? row[1].toString() : '';
      const tipoRegistro = row[2] != null ? row[2].toString() : '';
      const entrada = Number(row[3]) || 0;
      const salida = Number(row[4]) || 0;

      if (!saldos[id]) {
        saldos[id] = {
          nombre: nombre,
          tipos: {}
        };
      }

      if (!saldos[id].tipos[tipoRegistro]) {
        saldos[id].tipos[tipoRegistro] = {
          entradas: 0,
          salidas: 0
        };
      }

      saldos[id].tipos[tipoRegistro].entradas += entrada;
      saldos[id].tipos[tipoRegistro].salidas += salida;
    }

    // Convertir a array y filtrar solo saldos positivos
    const resultado = [];
    for (const [id, info] of Object.entries(saldos)) {
      for (const [tipo, montos] of Object.entries(info.tipos)) {
        const saldo = montos.salidas - montos.entradas;
        if (saldo > 0) {
          resultado.push({
            id: id,
            nombre: info.nombre,
            tipoRegistro: tipo,
            saldo: saldo
          });
        }
      }
    }

    // Si hay filtro, aplicar búsqueda parcial (case-insensitive) sobre id, nombre y tipoRegistro
    const aplicarFiltro = (item) => {
      if (!filtro && !nombreColaborador) return true;
      const q = (filtro || '').toString().trim().toLowerCase();
      const qName = (nombreColaborador || '').toString().trim().toLowerCase();

      let pasaFiltro = true;
      if (q) {
        pasaFiltro = (
          (item.id || '').toString().toLowerCase().indexOf(q) !== -1 ||
          (item.nombre || '').toString().toLowerCase().indexOf(q) !== -1 ||
          (item.tipoRegistro || '').toString().toLowerCase().indexOf(q) !== -1
        );
      }
      if (qName) {
        pasaFiltro = pasaFiltro && ((item.nombre || '').toString().toLowerCase().indexOf(qName) !== -1);
      }
      return pasaFiltro;
    };

    const filtrado = resultado.filter(aplicarFiltro);
    return filtrado;
  } catch (error) {
    console.error("Error en obtenerResumenSaldos:", error);
    return [];
  }
}

// ===================== Funciones auxiliares para vales y envío =====================

/**
 * Obtiene el email del colaborador desde la hoja `Colaboradores`.
 * Se asume que el email está en la columna I (índice 8, 0-based).
 * @param {string} idColaborador
 * @returns {string|null} email o null
 */
function obtenerEmailColaborador(idColaborador) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_COLABORADORES);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] != null && row[0].toString() === idColaborador.toString()) {
        // Columna I -> índice 8
        return row[8] ? row[8].toString().trim() : null;
      }
    }
    return null;
  } catch (e) {
    console.error('Error en obtenerEmailColaborador:', e);
    return null;
  }
}

/**
 * Obtiene un email de administrador desde la hoja `Usuarios` (primer admin encontrado).
 * @returns {string|null}
 */
function obtenerEmailAdmin() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_USUARIOS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const email = data[i][0] ? data[i][0].toString().trim() : '';
      const rol = data[i][1] ? data[i][1].toString().trim().toUpperCase() : '';
      if (rol === ROLES.ADMIN && email) return email;
    }
    // fallback: usuario activo
    return Session.getActiveUser().getEmail() || null;
  } catch (e) {
    console.error('Error en obtenerEmailAdmin:', e);
    return null;
  }
}

/**
 * Crea (si no existe) una carpeta llamada 'Vales_Caja' en la raíz del Drive del usuario/script.
 * @returns {Folder} carpeta creada o existente
 */
function crearCarpetaValesSiNoExiste() {
  try {
    const nombre = 'Vales_Caja';
    const folders = DriveApp.getFoldersByName(nombre);
    if (folders.hasNext()) {
      return folders.next();
    }
    const folder = DriveApp.createFolder(nombre);
    return folder;
  } catch (e) {
    console.error('Error en crearCarpetaValesSiNoExiste:', e);
    throw e;
  }
}

/**
 * Crea (si no existe) una carpeta llamada 'Informes de Caja' en la raíz del Drive.
 * @returns {Folder} carpeta creada o existente
 */
function crearCarpetaInformesSiNoExiste() {
  try {
    const nombre = 'Informes de Caja';
    const folders = DriveApp.getFoldersByName(nombre);
    if (folders.hasNext()) {
      return folders.next();
    }
    const folder = DriveApp.createFolder(nombre);
    return folder;
  } catch (e) {
    console.error('Error en crearCarpetaInformesSiNoExiste:', e);
    throw e;
  }
}

/**
 * Genera un PDF a partir de un Google Doc y lo guarda en la carpeta indicada.
 * @param {string} docId
 * @param {Folder} carpeta
 * @returns {object} { fileId, url }
 */
function guardarPdfDesdeDoc(docId, carpeta) {
  try {
    const fileDoc = DriveApp.getFileById(docId);
    const blobPdf = fileDoc.getAs(MimeType.PDF);
    const nombrePdf = fileDoc.getName() + '.pdf';
    const filePdf = carpeta.createFile(blobPdf).setName(nombrePdf);
    return { fileId: filePdf.getId(), url: filePdf.getUrl() };
  } catch (e) {
    console.error('Error en guardarPdfDesdeDoc:', e);
    throw e;
  }
}

/**
 * Comparte un archivo (Docs o PDF) con una lista de emails como viewers.
 * @param {string} fileId
 * @param {Array<string>} emails
 */
function compartirArchivoConEmails(fileId, emails) {
  try {
    const file = DriveApp.getFileById(fileId);
    emails.forEach(email => {
      try {
        file.addViewer(email);
      } catch (e) {
        console.warn('No se pudo compartir con', email, e);
      }
    });
  } catch (e) {
    console.error('Error en compartirArchivoConEmails:', e);
  }
}

/**
 * Envía un correo con el PDF adjunto (usando MailApp).
 * @param {string} emailTo
 * @param {string} subject
 * @param {string} body
 * @param {string} filePdfId
 */
function enviarCorreoConAdjunto(emailTo, subject, body, filePdfId) {
  try {
    if (!emailTo || !filePdfId) return;
    const filePdf = DriveApp.getFileById(filePdfId);
    const blob = filePdf.getAs(MimeType.PDF);
    MailApp.sendEmail({
      to: emailTo,
      subject: subject,
      body: body,
      attachments: [blob]
    });
  } catch (e) {
    console.error('Error en enviarCorreoConAdjunto:', e);
  }
}

// =============================================================================
// GESTIÓN DE PONDERACIÓN DE BONOS
// =============================================================================

/**
 * Obtiene la matriz de datos para la página de ponderación.
 * AJUSTE: Ahora carga los proyectos desde 'Bonificaciones', pero los valores de
 * ponderación vienen de la tabla estándar, con posibilidad de ser sobreescritos.
 * @returns {object} Objeto con { colaboradores, proyectos }.
 */
function obtenerDatosPonderacion() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetColaboradores = ss.getSheetByName(HOJA_COLABORADORES);
    const sheetBonificaciones = ss.getSheetByName(HOJA_BONIFICACIONES);
    const sheetPonderacion = ss.getSheetByName(HOJA_PONDERACION);

    // 1. Obtener colaboradores activos
    const colaboradoresData = sheetColaboradores.getDataRange().getValues().slice(1);
    const colaboradoresActivos = colaboradoresData
      .filter(row => row[6] === 'Activo')
      .map(row => ({ id: row[0].toString().trim(), nombre: row[1].toString().trim() }))
      .sort((a, b) => a.nombre.localeCompare(b.nombre));

    // 2. Obtener ponderaciones estándar (nuestra fuente principal de valores)
    const ponderacionesEstandarData = obtenerPonderacionEstandar(); // Usa la función que ya creamos
    const mapaPonderacionEstandar = new Map(ponderacionesEstandarData.map(p => [p.id.toString().trim(), p.ponderacion]));

    // 3. Obtener la lista de proyectos desde la hoja BONIFICACIONES
    if (!sheetBonificaciones) throw new Error("La hoja 'Bonificaciones' no existe. Por favor, calcule primero las bonificaciones.");
    const bonificacionesData = sheetBonificaciones.getDataRange().getValues();
    if (bonificacionesData.length <= 1) {
        // Si no hay datos, retornamos una estructura vacía para evitar errores en el frontend.
        return { colaboradores: colaboradoresActivos, proyectos: [] };
    }
    const proyectosEnBonificaciones = bonificacionesData.slice(1).map(row => row[0]).filter(String);

    // 4. Leer ponderaciones específicas (las que el admin guardó para un proyecto)
    const ponderacionesGuardadas = {};
    if (sheetPonderacion && sheetPonderacion.getLastRow() > 1) {
      const data = sheetPonderacion.getDataRange().getValues();
      const headers = data[0];
      const mapaNombreColumna = new Map(headers.map((h, i) => [h, i]));
      
      data.slice(1).forEach(row => {
        const proyecto = row[0];
        if (proyecto) {
          ponderacionesGuardadas[proyecto] = {};
          colaboradoresActivos.forEach(colab => {
            const colIndex = mapaNombreColumna.get(colab.nombre);
            if (colIndex !== undefined && row[colIndex] !== '') {
              ponderacionesGuardadas[proyecto][colab.id] = row[colIndex];
            }
          });
        }
      });
    }

    // 5. Construir la matriz final combinando los datos
    const proyectosResult = proyectosEnBonificaciones.map(nombreProyecto => {
      const ponderaciones = {};
      colaboradoresActivos.forEach(colab => {
        const valorGuardado = (ponderacionesGuardadas[nombreProyecto] && ponderacionesGuardadas[nombreProyecto][colab.id] !== undefined)
          ? ponderacionesGuardadas[nombreProyecto][colab.id]
          : null; // Si hay un valor específico guardado (incluso 0), lo usamos.

        const valorEstandar = mapaPonderacionEstandar.get(colab.id) || 0;
        
        // Prioridad: 1. Valor guardado específico. 2. Valor estándar.
        ponderaciones[colab.id] = (valorGuardado !== null) ? valorGuardado : valorEstandar;
      });
      return {
        nombre: nombreProyecto,
        ponderaciones: ponderaciones
      };
    });

    return {
      colaboradores: colaboradoresActivos,
      proyectos: proyectosResult
    };

  } catch (e) {
    console.error("Error en obtenerDatosPonderacion:", e);
    return { error: e.message };
  }
}

/**
 * Formatea un número como moneda CLP (Peso Chileno) en el lado del servidor.
 * @param {number} valor El número a formatear.
 * @returns {string} El valor formateado como "$1.234".
 */
function formatearMonedaCLP(valor) {
  if (typeof valor !== 'number') {
    return '$0';
  }
  // Añade el signo $ y los separadores de miles (puntos).
  return '$' + valor.toFixed(0).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1.');
}


/**
 * Guarda o actualiza una fila de ponderación para un proyecto específico.
 * @param {object} dataFila - Objeto con { nombreProyecto: '...', ponderaciones: { colabId: valor, ... } }.
 * @returns {object} Un objeto con { success: true/false, message: '...' }.
 */
function guardarPonderacionFila(dataFila) {
  try {
    const { nombreProyecto, ponderaciones } = dataFila;
    if (!nombreProyecto || !ponderaciones) {
      throw new Error("Datos incompletos. Se requiere nombre del proyecto y ponderaciones.");
    }

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    let sheet = ss.getSheetByName(HOJA_PONDERACION);
    if (!sheet) {
      sheet = ss.insertSheet(HOJA_PONDERACION);
      sheet.getRange(1, 1).setValue("Proyecto").setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
    }

    const sheetData = sheet.getDataRange().getValues();
    const headers = sheetData.length > 0 ? sheetData[0] : ["Proyecto"];
    
    // Obtener mapa de ID a Nombre para los headers
    const sheetColaboradores = ss.getSheetByName(HOJA_COLABORADORES);
    const colaboradoresData = sheetColaboradores.getDataRange().getValues().slice(1);
    const mapaIdANombre = new Map(colaboradoresData.map(row => [row[0].toString().trim(), row[1].toString().trim()]));

    // --- Sincronizar Cabeceras (Columnas) ---
    const mapaNombreAColumna = new Map(headers.map((h, i) => [h, i + 1]));
    let headerChanged = false;
    for (const colabId in ponderaciones) {
      const nombreColab = mapaIdANombre.get(colabId);
      if (nombreColab && !mapaNombreAColumna.has(nombreColab)) {
        const newColIndex = sheet.getLastColumn() + 1;
        sheet.getRange(1, newColIndex).setValue(nombreColab).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
        mapaNombreAColumna.set(nombreColab, newColIndex);
        headerChanged = true;
      }
    }
    // Si se añadieron columnas, redimensionar
    if(headerChanged) sheet.autoResizeColumns(1, sheet.getLastColumn());


    // --- Preparar la Fila de Datos ---
    const filaParaGuardar = new Array(mapaNombreAColumna.size).fill(0);
    filaParaGuardar[0] = nombreProyecto;

    for (const colabId in ponderaciones) {
        const nombreColab = mapaIdANombre.get(colabId);
        if (nombreColab && mapaNombreAColumna.has(nombreColab)) {
            const colIdx = mapaNombreAColumna.get(nombreColab);
            filaParaGuardar[colIdx - 1] = ponderaciones[colabId] || 0;
        }
    }
    
    // --- Encontrar y Actualizar/Insertar Fila ---
    const proyectosEnHoja = sheet.getRange(2, 1, sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 1).getValues().flat();
    const rowIndex = proyectosEnHoja.findIndex(p => p === nombreProyecto);

    if (rowIndex !== -1 && proyectosEnHoja[0] !== '') {
      // Actualizar fila existente
      const rowNum = rowIndex + 2; // +1 por 0-based, +1 por header
      sheet.getRange(rowNum, 1, 1, filaParaGuardar.length).setValues([filaParaGuardar]);
    } else {
      // Añadir nueva fila
      sheet.appendRow(filaParaGuardar);
    }

    return { success: true, message: "Ponderación guardada correctamente." };
  } catch (e) {
    console.error("Error en guardarPonderacionFila:", e);
    return { success: false, message: `Error al guardar: ${e.message}` };
  }
}

/**
 * Guarda los resultados del cálculo de bonos en una nueva hoja llamada 'Bonos a Pagar'.
 * @param {Array<Array<string>>} dataTabla - Un array 2D con los datos de la tabla de resultados.
 * @returns {object} Un objeto con { success: true/false, message: '...' }.
 */
function guardarBonosAPagar(dataTabla) {
  try {
    if (!dataTabla || dataTabla.length === 0) {
      throw new Error("No se recibieron datos para guardar.");
    }

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const nombreHoja = "Bonos a Pagar";
    let sheet = ss.getSheetByName(nombreHoja);

    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet(nombreHoja);
    }

    const numRows = dataTabla.length;
    const numCols = dataTabla[0].length;

    // Escribir todos los datos de una vez
    sheet.getRange(1, 1, numRows, numCols).setValues(dataTabla);

    // --- Aplicar Formato ---
    // Formato de cabecera
    sheet.getRange(1, 1, 1, numCols).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
    // Formato de fila de totales
    sheet.getRange(numRows, 1, 1, numCols).setBackground("#f0f0f0").setFontWeight("bold");
    // Formato de columna de totales
     sheet.getRange(1, numCols, numRows, 1).setFontWeight("bold");

    // Formato de moneda para los valores numéricos (evitando cabeceras y primera columna)
    sheet.getRange(2, 2, numRows - 1, numCols - 1).setNumberFormat('$#,##0.00');
    
    sheet.autoResizeColumns(1, numCols);

    return { success: true, message: `Resultados guardados correctamente en la hoja '${nombreHoja}'.` };
  } catch (e) {
    console.error("Error en guardarBonosAPagar:", e);
    return { success: false, message: `Error al guardar en la hoja de cálculo: ${e.message}` };
  }
}

function obtenerBalanceFiltrado(filtros) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(HOJA_CONTABILIDAD);
    const data = sheet.getDataRange().getValues().slice(1);

    const fechaDesde = new Date(filtros.fechaDesde.replace(/-/g, '\/') + ' 00:00:00');
    const fechaHasta = new Date(filtros.fechaHasta.replace(/-/g, '\/') + ' 23:59:59');

    let saldoAcumulado = 0;
    let totalEntradas = 0;
    let totalSalidas = 0;

    const resultados = data.map(row => {
        const fechaMovimiento = new Date(row[6]);
        const tipoRegistro = row[2];
        const entrada = parseFloat(row[3]) || 0;
        const salida = parseFloat(row[4]) || 0;

        // Lógica de filtrado
        const enRango = fechaMovimiento >= fechaDesde && fechaMovimiento <= fechaHasta;
        const coincideTipo = (filtros.tipoMovimiento === 'TODOS' || tipoRegistro === filtros.tipoMovimiento);

        if (enRango && coincideTipo) {
            saldoAcumulado += entrada - salida;
            totalEntradas += entrada;
            totalSalidas += salida;
            return {
                fecha: Utilities.formatDate(fechaMovimiento, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
                nombre: row[1],
                entrada: entrada,
                salida: salida,
                saldo: saldoAcumulado,
                detalle: row[5] || ''
            };
        }
        return null;
    }).filter(Boolean); // Elimina los nulos que no pasaron el filtro

    return {
      data: resultados,
      summary: {
        totalEntradas: totalEntradas,
        totalSalidas: totalSalidas,
        diferencia: totalEntradas - totalSalidas
      }
    };

  } catch (e) {
    console.error("Error en obtenerBalanceFiltrado:", e);
    return { error: e.message };
  }
}

function exportarBalanceComoNuevaHoja(datosBalance) {
  try {
    // 1. Generar nombre de archivo con fecha y hora
    const ahora = new Date();
    const fechaFormateada = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
    const nombreArchivo = `Informe de Balance - ${fechaFormateada}`;

    // 2. Crear la nueva hoja de cálculo
    const newSs = SpreadsheetApp.create(nombreArchivo);
    const newFile = DriveApp.getFileById(newSs.getId());

    // 3. Obtener/Crear la carpeta de destino
    const carpetaDestino = crearCarpetaInformesSiNoExiste();

    // 4. Mover el archivo a la carpeta
    newFile.moveTo(carpetaDestino);

    // 5. Llenar la hoja con los datos
    guardarBalanceEnHoja({ ...datosBalance, spreadsheet: newSs });

    // 6. Devolver la URL del archivo ya movido
    return { success: true, url: newSs.getUrl() };

  } catch (e) {
    console.error("Error en exportarBalanceComoNuevaHoja:", e);
    return { success: false, message: `Error al exportar: ${e.message}` };
  }
}

function guardarBalanceEnHoja(payload) {
    const datosBalance = payload;
    const ss = payload.spreadsheet || SpreadsheetApp.openById(getSpreadsheetId());

    try {
        const nombreHoja = payload.spreadsheet ? "Balance" : "Balance Filtrado";
        let sheet = ss.getSheetByName(nombreHoja);
        if (sheet && !payload.spreadsheet) {
          sheet.clear();
        } else if (!sheet) {
          sheet = ss.insertSheet(nombreHoja);
        }

        // 1. Añade "Detalle" al final de los encabezados
        const headers = ["Fecha", "Colaborador", "Entrada", "Salida", "Saldo Acumulado", "Detalle"];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers])
          .setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");

        if (datosBalance && datosBalance.data && datosBalance.data.length > 0) {
          // 2. Añade r.detalle al mapeo de las filas
          const rows = datosBalance.data.map(r => [r.fecha, r.nombre, r.entrada, r.salida, r.saldo, r.detalle]);
          sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

          // La lógica del pie de página no cambia
          const summaryRow = sheet.getLastRow() + 2;
          sheet.getRange(summaryRow, 2).setValue("TOTALES").setFontWeight("bold");
          sheet.getRange(summaryRow, 3).setValue(datosBalance.summary.totalEntradas);
          sheet.getRange(summaryRow, 4).setValue(datosBalance.summary.totalSalidas);

          sheet.getRange(summaryRow + 1, 2).setValue("DIFERENCIA").setFontWeight("bold");
          sheet.getRange(summaryRow + 1, 5).setValue(datosBalance.summary.diferencia);

          sheet.getRange(2, 3, sheet.getLastRow(), 3).setNumberFormat("$#,##0");
        }

        sheet.autoResizeColumns(1, headers.length);
        return { success: true, message: `Datos guardados en la hoja '${nombreHoja}'.` };
    } catch(e) {
        return { success: false, message: `Error al guardar: ${e.message}` };
    }
}