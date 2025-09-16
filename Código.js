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

const ROLES = {
  ADMIN: "ADMINISTRADOR",
  ASISTENTE: "ASISTENTE",
  SIN_ACCESO: "SIN_ACCESO"
};

// ===============================================================
// SERVIDOR WEB Y AUTENTICACIÓN
// ===============================================================

function doGet() {
  // Crea una plantilla a partir de index.html y la evalúa para procesar las etiquetas <?!= ... ?>
  return HtmlService.createTemplateFromFile('index.html')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle("Sistema de Gestión de Nómina");
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

    // Obtener lista de proyectos
    const proyectos = sheetProyectos ? sheetProyectos.getRange("B2:B").getValues().flat().filter(String) : [];

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
      const headers = [["project_code", "project_name", "registration_date", "project_address", "project_georeference", "project_contact", "project_observation", "Timestamp"]];
      sheetProyectos.getRange(1, 1, 1, 8).setValues(headers).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
      sheetProyectos.getRange("A:A").setNumberFormat("@");
      sheetProyectos.getRange("C:C").setNumberFormat("dd/mm/yyyy");
      sheetProyectos.getRange("H:H").setNumberFormat("dd/mm/yyyy hh:mm:ss");
      sheetProyectos.autoResizeColumns(1, 8);
    }
    
    return "Sistema inicializado correctamente. Todas las hojas han sido creadas y configuradas.";
  } catch (error) {
    console.error("Error en inicializarSistema:", error);
    return `Error al inicializar: ${error.message}`;
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
      project_address: row[3] || "",
      project_georeference: row[4] || "",
      project_contact: row[5] || "",
      project_observation: row[6] || ""
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
        sheet.getRange(i + 1, 4).setValue(proyecto.project_address || "");
        sheet.getRange(i + 1, 5).setValue(proyecto.project_georeference || "");
        sheet.getRange(i + 1, 6).setValue(proyecto.project_contact || "");
        sheet.getRange(i + 1, 7).setValue(proyecto.project_observation || "");
        
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