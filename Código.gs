/**
 * SIEL SERVER CORE v8.0 - Master
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SIEL - Sistema Informático de Laboratorio')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 1. Carga inicial de datos (Exámenes y Perfiles)
function getInitData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DATOS DE EXÁMENES
    const sheetEx = ss.getSheetByName("BD_NORMALIZADA"); 
    if (!sheetEx) throw new Error("Pestaña 'BD_NORMALIZADA' no encontrada.");
    const dataEx = sheetEx.getDataRange().getDisplayValues();
    const headersEx = dataEx[0].map(h => h.trim());
    const examenes = dataEx.slice(1).map(row => {
      let obj = {}; headersEx.forEach((h, i) => { obj[h] = row[i]; }); return obj;
    });

    // DATOS DE PERFILES
    let perfiles = [];
    const sheetPerf = ss.getSheetByName("PERFILES");
    if (sheetPerf) {
      const dataPerf = sheetPerf.getDataRange().getDisplayValues();
      if(dataPerf.length > 1) {
        const headersPerf = dataPerf[0].map(h => h.trim());
        perfiles = dataPerf.slice(1).map(row => {
          let obj = {}; headersPerf.forEach((h, i) => { obj[h] = row[i]; }); return obj;
        });
      }
    }

    return JSON.stringify({ examenes: examenes, perfiles: perfiles });
  } catch (e) { throw new Error(e.message); }
}

// 2. Validación de Usuarios
function verificarLogin(emailIngresado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUsuarios = ss.getSheetByName("USUARIOS");
    
    if (!sheetUsuarios) {
      return { success: true, role: "ADMIN", msg: "Hoja USUARIOS no existe. Modo Admin activo." };
    }

    const data = sheetUsuarios.getDataRange().getValues();
    const emailLower = emailIngresado.trim().toLowerCase();
    
    // Buscar en la hoja usuarios (Columna A: Email, Columna B: Rol)
    for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim().toLowerCase() === emailLower) {
            return { success: true, role: data[i][1].toString().toUpperCase().trim() };
        }
    }
    
    return { success: false, msg: "Usuario no registrado. Contacte al administrador." };
  } catch (e) {
    return { success: false, msg: "Error del servidor: " + e.message };
  }
}

// 3. Guardado Masivo
function guardarEdicionMasiva(datosModificados) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_NORMALIZADA");
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.trim());
    const numIdx = headers.indexOf("Número");
    
    const modMap = {};
    datosModificados.forEach(d => { modMap[d['Número']] = d; });

    for (let i = 1; i < data.length; i++) {
      let id = String(data[i][numIdx]);
      if (modMap[id]) {
        headers.forEach((h, colIdx) => {
          if (modMap[id][h] !== undefined) data[i][colIdx] = modMap[id][h];
        });
      }
    }
    sheet.getDataRange().setValues(data);
    return "OK";
  } catch (e) { throw new Error(e.message); }
}
