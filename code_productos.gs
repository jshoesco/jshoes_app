// =============================================================================
// 1. CONFIGURACIÓN Y SETUP DE INTERFAZ
// =============================================================================

// IDs de recursos externos
const INPUTS_SS_ID = "1f-fDYQUPD8S8M-KevcvCJm57s93OI1aC9alINnHlnds"; // ID de la Hoja de Inputs
const DRIVE_FOLDER_ID = "1FOwQiJNgD4go2dLricC3ZdRf3LbrH3w5"; // ID de la Carpeta de Drive

// ⚡ ¡NUEVO! Define quién puede usar la App Web.
// ⚠️ ¡CAMBIA ESTO POR TU EMAIL!
const ADMIN_EMAIL = "tennisandariegos@gmail.com"; // ✅ CONFIRMADO

/**
 * Crea un menú personalizado en la hoja de cálculo al abrir.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Catálogo")
    .addItem("Registrar producto", "abrirForm")
    .addSeparator()
    .addItem("⚡ (Temporal) Llenar IDs de Imágenes", "temp_llenarIdsDeImagenes")
    .addToUi();
}

/**
 * Muestra el formulario como una barra lateral (sidebar).
 */
function abrirForm() {
  const html = HtmlService.createHtmlOutputFromFile('form')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * ⚡ ¡ACTUALIZADO CON SEGURIDAD!
 * Sirve el formulario como una aplicación web SEGURA.
 */
function doGet(e) {
  var email = Session.getActiveUser().getEmail();
  
  if (email.toLowerCase() === ADMIN_EMAIL.toLowerCase()) {
    const html = HtmlService.createHtmlOutputFromFile('form')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("Registro de Catálogo"); 
    return html;
  } else {
    return HtmlService.createHtmlOutput(
      `<h1>Acceso Denegado</h1><p>No tienes permiso para ver esta aplicación.</p>`
    ).setTitle("Acceso Denegado");
  }
}

// =============================================================================
// 2. OPERACIONES DE DATOS (Lectura y Consultas)
// =============================================================================

/**
 * ⚡ ¡ACTUALIZADO!
 * La lista de 'Género' ahora está definida internamente (hard-coded).
 */
function getDropdownData() {
  const sh = SpreadsheetApp.getActive().getSheetByName("data");
  const data = sh.getDataRange().getValues();
  if (!data || data.length === 0) return {};
  const headers = data[0];
  const values = {};
  const modelosPorMarca = {};
  
  // ⚡ ¡CAMBIO! Se quitó "Género" de esta lista.
  const dropdownHeaders = ["Marca", "Modelo", "Proveedor", "Tipo"];
  
  headers.forEach((h, i) => {
    if (dropdownHeaders.includes(h)) {
      const colVals = data.slice(1).map(r => r[i]).filter(v => v !== "" && v != null);
      values[h] = [...new Set(colVals)].sort();
    }
  });

  // ⚡ ¡NUEVO! Añadimos la lista de Género manualmente en el orden solicitado.
  values["Género"] = ["Hombre", "Mujer", "Unisex", "Niños"];

  const marcaIndex = headers.indexOf("Marca");
  const modeloIndex = headers.indexOf("Modelo");
  if (marcaIndex !== -1 && modeloIndex !== -1) {
    data.slice(1).forEach(r => {
      const marca = r[marcaIndex];
      const modelo = r[modeloIndex];
      if (!marca || !modelo) return;
      if (!modelosPorMarca[marca]) modelosPorMarca[marca] = [];
      if (!modelosPorMarca[marca].includes(modelo)) modelosPorMarca[marca].push(modelo);
    });
    Object.keys(modelosPorMarca).forEach(marca => modelosPorMarca[marca].sort());
  }
  values["ModelosPorMarca"] = modelosPorMarca;
  return values;
}
function getUltimoSKU(proveedor) {
  const sh = SpreadsheetApp.getActive().getSheetByName("productos");
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const skuColIndex = headers.indexOf("Sku");
  const numFilas = sh.getLastRow() - 1;
  const dataSh = SpreadsheetApp.getActive().getSheetByName("data");
  const data = dataSh.getDataRange().getValues();
  const headers2 = data[0] || [];
  const provIdx = headers2.indexOf("Proveedor");
  const prefIdx = headers2.indexOf("PrefijoSku");
  let pref = "";
  const proveedorLimpio = proveedor ? proveedor.toString().trim().toLowerCase() : "";
  if (provIdx !== -1 && prefIdx !== -1 && proveedorLimpio) {
    const row = data.slice(1).find(r => r[provIdx] && r[provIdx].toString().trim().toLowerCase() === proveedorLimpio);
    if (row) pref = row[prefIdx] || "";
  }
  let ultimoNumero = 0;
  if (skuColIndex !== -1 && numFilas > 0 && pref) {
    const skus = sh.getRange(2, skuColIndex + 1, numFilas).getValues().flat();
    const numeros = skus.map(s => {
      if (typeof s !== "string") s = String(s || "");
      if (s.startsWith(pref)) {
        const match = s.match(/\d+$/);
        return match ? Number(match[0]) : null;
      }
      return null;
    }).filter(v => v != null && !isNaN(v));
    ultimoNumero = numeros.length ? Math.max(...numeros) : 0;
  }
  return { prefijo: pref, ultimo: ultimoNumero };
}
function getModeloInfo(modelo) {
  const ss = SpreadsheetApp.getActive();
  const modeloLimpio = modelo ? modelo.toString().toLowerCase() : "";
  const shData = ss.getSheetByName("data");
  const data = shData.getDataRange().getValues();
  const modeloIdx = data[0].indexOf("Modelo");
  const tipoIdx = data[0].indexOf("Tipo");
  let tipo = "";
  if (modeloIdx !== -1 && tipoIdx !== -1 && modeloLimpio) {
    const row = data.slice(1).find(r => r[modeloIdx] && r[modeloIdx].toString().toLowerCase() === modeloLimpio);
    if (row) tipo = row[tipoIdx] || "";
  }
  const shProductos = ss.getSheetByName("productos");
  let conteo = 0;
  if (shProductos.getLastRow() > 1 && modeloLimpio) {
    const headers = shProductos.getRange(1, 1, 1, shProductos.getLastColumn()).getValues()[0];
    const modeloColIndex = headers.indexOf("Modelo");
    if (modeloColIndex !== -1) {
      const dataProductos = shProductos.getRange(2, modeloColIndex + 1, shProductos.getLastRow() - 1, 1).getValues();
      conteo = dataProductos.filter(r => r[0] && r[0].toString().toLowerCase() === modeloLimpio).length;
    }
  }
  return { conteo, tipo };
}
function getTarifaAndMargen() {
  const ss = SpreadsheetApp.openById(INPUTS_SS_ID);
  const sh = ss.getSheetByName("inputs");
  const data = sh.getDataRange().getValues();
  const headers = data[0] || [];
  let margen = 0;
  try {
    margen = Number(ss.getRangeByName("margenGanancia").getValue()) || 0;
  } catch (e) {
    Logger.log("Error al obtener margenGanancia: " + e.message);
  }
  const idx = headers.indexOf("Tarifa");
  let tarifa = 0;
  if (idx !== -1 && data.length > 1) {
    const tarifas = data.slice(1).map(r => Number(r[idx])).filter(v => !isNaN(v));
    tarifa = tarifas.length ? Math.max(...tarifas) : 0;
  }
  return { tarifa, margen };
}
function getFolderUrl() {
  try {
    DriveApp.getFolderById(DRIVE_FOLDER_ID); 
    return DRIVE_FOLDER_ID;
  } catch (e) {
    Logger.log("Error al obtener ID de la carpeta: " + e.message);
    throw new Error("No se pudo acceder a la carpeta de imágenes. Verifique la constante DRIVE_FOLDER_ID.");
  }
}

// =============================================================================
// 3. ⚡ ¡ESTRUCTURA DE REGISTRO CON LOCKSERVICE! (TRANSACCIÓN + ROBUSTEZ)
// =============================================================================

/**
 * ⚡ ¡ACTUALIZADO!
 * Función "Maestra" que ahora es 100% robusta contra duplicados.
 * Acepta un 'oldSkuToReplace' opcional para marcar un producto antiguo como reemplazado.
 */
function registrarProductoYSubirImagen(d, imgDataURL, oldSkuToReplace = null) {
  // 1. Solicitar el bloqueo del script.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 15 segundos de timeout
  } catch (e) { {
    throw new Error("El servidor está ocupado procesando otro registro. Por favor, intenta de nuevo en 15 segundos.");
  }}
  
  let fileId = null; // Variable para rollback
  
  try {
    // --- SECCIÓN CRÍTICA (DENTRO DEL BLOQUEO) ---
    
    // Paso 1: (NUEVO) Desactivar el Sku antiguo si se proporcionó
    if (oldSkuToReplace) {
      _desactivarSkuAntiguo(oldSkuToReplace);
    }

    // Paso 2: Generar el SKU final y seguro.
    const skuInfo = _getUltimoSKU_interno(d.proveedor); 
    const finalSku = `${skuInfo.prefijo || ''}${String(skuInfo.ultimo + 1).padStart(3,"0")}`;
    d.sku = finalSku; // Asignar el SKU oficial al objeto de datos

    // Paso 3: Generar el Nombre de Imagen final y seguro
    const fecha = d.fecha;
    const fechaYYMMDD = fecha.slice(2,4) + fecha.slice(5,7) + fecha.slice(8,10);
    const finalNombreImagen = `${fechaYYMMDD}-${d.sku}-${d.marca}-${d.modelo}-${d.tipo}-${d.genero}`.toLowerCase().replace(/[^a-z0-9\-]/g,"_") + ".jpg";
    d.nombreImagen = finalNombreImagen;

    // Paso 4: Subir la imagen a Drive
    fileId = _subirImagenInterna(imgDataURL, d.nombreImagen);
    d.idImagen = fileId;
    
    // Paso 5: Escribir los datos en la hoja de cálculo
    _registrarProductoEnHoja(d); // Esta función ahora asigna "Activo" al estado
    
    // --- FIN DE LA SECCIÓN CRÍTICA ---
    
    lock.releaseLock(); // Liberar el bloqueo
    
    // Devolver los datos generados para que el cliente los muestre
    return {
      status: 'success',
      sku: d.sku,
      nombreImagen: d.nombreImagen
    };
    
  } catch (e) {
    // ¡ROLLBACK! Si algo falló, borramos la imagen que se subió.
    if (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } 
      catch (trashError) { Logger.log(`Error al intentar borrar imagen huérfana ${fileId}: ${trashError.message}`); }
    }
    
    lock.releaseLock(); // Asegurarse de liberar el bloqueo en caso de error
    throw new Error(`Error en la transacción: ${e.message}`);
  }
}

/**
 * ⚡ ¡NUEVA FUNCIÓN HELPER!
 * Busca un Sku, borra su imagen de Drive y lo marca como "Reemplazado".
 * ⚡ ¡ACTUALIZADO PARA USAR COLUMNA N (Status)!
 */
function _desactivarSkuAntiguo(oldSku) {
  if (!oldSku) return;
  
  const sh = SpreadsheetApp.getActive().getSheetByName("productos");
  if (!sh) throw new Error("Hoja 'productos' no encontrada.");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const skuColIdx = headers.findIndex(h => h.toLowerCase().trim() === 'sku');
  const idImgColIdx = headers.findIndex(h => h.toLowerCase().trim() === 'id imagen');
  const nombreImgColIdx = headers.findIndex(h => h.toLowerCase().trim() === 'nombre img');
  // ⚡ ¡CAMBIO! Buscando "Status" (Col N)
  const statusColIdx = headers.findIndex(h => h.toLowerCase().trim() === 'status'); 

  if (skuColIdx === -1) throw new Error("No se encontró la columna 'Sku'.");
  // ⚡ ¡CAMBIO! Actualizado el mensaje de error
  if (statusColIdx === -1) throw new Error("No se encontró la columna 'Status' (Columna N). Por favor, añádela a la hoja 'productos'.");
  if (idImgColIdx === -1) throw new Error("No se encontró la columna 'ID Imagen'.");

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  let foundRowIndex = -1;
  let fileIdToDelete = null;

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][skuColIdx]).trim().toUpperCase() === String(oldSku).trim().toUpperCase()) {
      foundRowIndex = i + 2; // +2 porque el índice es base 0 y la data empieza en fila 2
      fileIdToDelete = data[i][idImgColIdx];
      break;
    }
  }

  if (foundRowIndex > -1) {
    // 1. Marcar como Reemplazado en la Col N "Status"
    sh.getRange(foundRowIndex, statusColIdx + 1).setValue("Reemplazado");
    
    // 2. Borrar la imagen de Drive si existe
    if (fileIdToDelete) {
      try {
        DriveApp.getFileById(fileIdToDelete).setTrashed(true);
        // 3. Limpiar los campos de imagen en la hoja
        sh.getRange(foundRowIndex, idImgColIdx + 1).setValue("");
        if (nombreImgColIdx > -1) {
          sh.getRange(foundRowIndex, nombreImgColIdx + 1).setValue("");
        }
      } catch (e) {
        Logger.log(`No se pudo borrar la imagen ${fileIdToDelete} del Sku ${oldSku}. Puede que ya estuviera borrada. Error: ${e.message}`);
        // Continuamos incluso si la imagen no se puede borrar (podría ya no existir)
      }
    }
  } else {
    // Si no se encuentra, lanzamos un error para detener la transacción
    throw new Error(`El Sku a reemplazar (${oldSku}) no fue encontrado. Verifica el Sku e intenta de nuevo.`);
  }
}


/**
 * ⚡ ¡NUEVO HELPER INTERNO!
 * Esta es la versión 'getUltimoSKU' que SÓLO usa el servidor.
 */
function _getUltimoSKU_interno(proveedor) {
  const sh = SpreadsheetApp.getActive().getSheetByName("productos");
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const skuColIndex = headers.indexOf("Sku");
  const numFilas = sh.getLastRow() - 1;
  const dataSh = SpreadsheetApp.getActive().getSheetByName("data");
  const data = dataSh.getDataRange().getValues();
  const headers2 = data[0] || [];
  const provIdx = headers2.indexOf("Proveedor");
  const prefIdx = headers2.indexOf("PrefijoSku");
  let pref = "";
  const proveedorLimpio = proveedor ? proveedor.toString().trim().toLowerCase() : "";
  if (provIdx !== -1 && prefIdx !== -1 && proveedorLimpio) {
    const row = data.slice(1).find(r => r[provIdx] && r[provIdx].toString().trim().toLowerCase() === proveedorLimpio);
    if (row) pref = row[prefIdx] || "";
  }
  let ultimoNumero = 0;
  if (skuColIndex !== -1 && numFilas > 0 && pref) {
    const skus = sh.getRange(2, skuColIndex + 1, numFilas).getValues().flat();
    const numeros = skus.map(s => {
      if (typeof s !== "string") s = String(s || "");
      if (s.startsWith(pref)) {
        const match = s.match(/\d+$/);
        return match ? Number(match[0]) : null;
      }
      return null;
    }).filter(v => v != null && !isNaN(v));
    ultimoNumero = numeros.length ? Math.max(...numeros) : 0;
  }
  return { prefijo: pref, ultimo: ultimoNumero };
}


/**
 * Lógica de subida de imagen (interna).
 */
function _subirImagenInterna(imgDataURL, nombreArchivo) {
  if (!imgDataURL) throw new Error("No se recibió data de imagen.");
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const base64 = imgDataURL.match(/base64,(.*)$/)[1];
  const contentType = imgDataURL.match(/data:(.*);base64,/)[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), contentType, nombreArchivo);
  const file = folder.createFile(blob);
  return file.getId(); 
}

/**
 * ⚡ ¡ACTUALIZADO!
 * Lógica de escritura en la hoja (interna).
 * Ahora asigna "Activo" a la columna "Status" (Col N) por defecto.
 */
function _registrarProductoEnHoja(d) {
  const sh = SpreadsheetApp.getActive().getSheetByName("productos");
  if (!sh) throw new Error("Hoja 'productos' no encontrada.");
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const newRow = new Array(headers.length);

  headers.forEach((header, index) => {
    const headerNorm = header.toLowerCase().trim(); 
    if (headerNorm === "fecha") newRow[index] = d.fecha;
    else if (headerNorm === "proveedor") newRow[index] = d.proveedor;
    else if (headerNorm === "sku") newRow[index] = d.sku;
    else if (headerNorm === "nombre") newRow[index] = d.nombre;
    else if (headerNorm === "marca") newRow[index] = d.marca;
    else if (headerNorm === "modelo") newRow[index] = d.modelo;
    else if (headerNorm === "género" || headerNorm === "genero") newRow[index] = d.genero; 
    else if (headerNorm === "tipo") newRow[index] = d.tipo;
    else if (headerNorm === "costo") newRow[index] = d.costo;
    else if (headerNorm === "ganancia") newRow[index] = d.ganancia;
    else if (headerNorm === "precio") newRow[index] = d.precio;
    else if (headerNorm === "nombre img") newRow[index] = d.nombreImagen;
    else if (headerNorm === "id imagen") newRow[index] = d.idImagen;
    // ⚡ ¡CAMBIO AQUÍ!
    else if (headerNorm === "status") newRow[index] = "Activo"; // Asigna "Activo" a la nueva columna "Status"
    else {
      // Dejar vacío si no se mapea
      newRow[index] = "";
    }
  });
  sh.appendRow(newRow);
  return true;
}


// ========================================================================================
// 5. FUNCIÓN TEMPORAL DE MIGRACIÓN
// ========================================================================================
// (La dejamos aquí por si acaso)
function temp_llenarIdsDeImagenes() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Iniciando Migración de IDs...", "Esto puede tardar varios minutos. No cierres la hoja. Se te notificará al terminar.", ui.ButtonSet.OK);
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("productos");
    if (!sheet) throw new Error("Hoja 'productos' no encontrada.");
    const range = sheet.getDataRange();
    const values = range.getValues(); 
    const headers = values[0];
    const nameCol = headers.findIndex(h => h.toLowerCase().trim() === 'nombre img');
    const idCol = headers.findIndex(h => h.toLowerCase().trim() === 'id imagen');
    if (nameCol === -1 || idCol === -1) {
      throw new Error("No se encontraron las columnas 'Nombre img' o 'ID Imagen'. Verifica las cabeceras de la Fila 1.");
    }
    ui.alert("Paso 1 de 3: Creando índice de imágenes de Drive... (Esto puede tardar uno o dos minutos)");
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const files = folder.getFiles();
    const fileIndex = new Map(); 
    while (files.hasNext()) {
      let file = files.next();
      fileIndex.set(file.getName(), file.getId());
    }
    let filesIndexed = fileIndex.size;
    if (filesIndexed === 0) throw new Error("No se encontraron archivos en la carpeta de Drive. Verifica la constante DRIVE_FOLDER_ID.");
    ui.alert(`Paso 2 de 3: Índice creado (${filesIndexed} imágenes). Comparando con ${values.length - 1} filas de la hoja...`);
    let idsToUpdate = 0;
    let errors = 0;
    for (let i = 1; i < values.length; i++) {
      let row = values[i];
      let fileName = row[nameCol];
      let fileId = row[idCol];
      if (fileName && !fileId) {
        let foundId = fileIndex.get(String(fileName).trim());
        if (foundId) {
          values[i][idCol] = foundId; 
          idsToUpdate++;
        } else {
          values[i][idCol] = "ERROR_NO_ENCONTRADO";
          errors++;
        }
      }
    }
    ui.alert(`Paso 3 de 3: Escribiendo ${idsToUpdate} IDs nuevos en la hoja...`);
    range.setValues(values); 
    let message = `¡Proceso completado! 🎉\n\nIDs de imagen encontrados y actualizados: ${idsToUpdate}`;
    if (errors > 0) {
      message += `\n\nImágenes no encontradas en Drive: ${errors}\n(Se marcaron como "ERROR_NO_ENCONTRADO". Búscalas y corrígelas manualmente.)`;
    }
    ui.alert(message);
  } catch (e) {
    Logger.log(e);
    ui.alert("Error en el proceso", e.message, ui.ButtonSet.OK);
  }
}
