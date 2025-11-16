// 🚨 ID DE TU LIBRO "PRODUCTOS" (CONFIRMADO)
var PRODUCTOS_SHEET_ID = "1pbjnxZtYwnnvBWQbbzdLc2XCEO9qaXTyOjne8PNgz8I"; 
// Nombre de la hoja dentro de ese libro.
var SHEET_NAME = "productos"; 
// ⚡ Llave para guardar el progreso del usuario
var USER_PROP_KEY = 'CONTROL_COMPARTIDAS_LAST_INDEX';

// =======================================================
// FUNCIONES AUXILIARES
// =======================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('form')
      .setTitle('Control de Compartidas')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSourceSheet() {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`Hoja "${SHEET_NAME}" no encontrada.`);
    }
    return sheet;
  } catch (e) {
    throw new Error('Error al abrir el libro "PRODUCTOS": ' + e.message);
  }
}

function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  // ⚡ ¡ACTUALIZADO! Mapeamos más campos para la edición
  headers.forEach((header, i) => {
    const norm = header.toLowerCase().trim();
    if (norm === "sku") map.sku = i;
    else if (norm === "marca") map.marca = i;
    else if (norm === "modelo") map.modelo = i;
    else if (norm === "género" || norm === "genero") map.genero = i;
    else if (norm === "precio") map.precio = i;
    else if (norm === "id imagen") map.idImagen = i; 
    else if (norm === "compartido") map.compartido = i; 
  });
  if (map.idImagen === undefined || map.compartido === undefined || map.sku === undefined) {
    throw new Error('No se encontraron columnas requeridas. Asegúrate de que existan las cabeceras "Sku", "ID Imagen" y "Compartido".');
  }
  return map;
}

function formatDescription(product) {
  var priceFormatted = '0'; 
  if (product.precio) {
    var priceValue = product.precio;
    priceFormatted = Math.round(priceValue).toString();
    priceFormatted = priceFormatted.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  }
  var description = 
    "SKU: " + product.sku + "\n" +
    "Marca: " + product.marca + "\n" +
    "Modelo: " + product.modelo + "\n" +
    "Género: " + product.genero + "\n" +
    "Precio: $" + priceFormatted; 
  return description;
}

// =======================================================
// FUNCIONES DE PROGRESO (Guardado)
// =======================================================

/**
 * Obtiene la última fila guardada para este usuario.
 */
function getUserProgress() {
  return PropertiesService.getUserProperties().getProperty(USER_PROP_KEY);
}

/**
 * Guarda la fila actual para este usuario.
 */
function saveUserProgress(rowIndex) {
  PropertiesService.getUserProperties().setProperty(USER_PROP_KEY, rowIndex);
}

/**
 * ⚡ ¡NUEVO!
 * Borra el progreso guardado del usuario.
 */
function resetUserProgress() {
  try {
    PropertiesService.getUserProperties().deleteProperty(USER_PROP_KEY);
    return { status: "success", message: "Progreso reseteado." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}


// =======================================================
// FUNCIONES DE LÓGICA PRINCIPAL
// =======================================================

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('✅ Control Compartidas')
      .addItem('Abrir Panel', 'showModal') 
      .addToUi();
}

function showModal() {
  var html = HtmlService.createHtmlOutputFromFile('form')
      .setTitle('Control de Compartidas')
      .setWidth(500) 
      .setHeight(650); 
  SpreadsheetApp.getUi().showModalDialog(html, 'Control de Compartidas');
}


function getNextProduct(indexToStartFrom) {
  var sheet;
  var map;
  var values;
  try {
    sheet = getSourceSheet();
    map = getHeaderMap(sheet); 
    values = sheet.getDataRange().getValues(); 
  } catch (e) {
    return { error: e.message }; 
  }
  
  var totalProducts = values.length - 1; 
  var startIndex = (indexToStartFrom < 1 || indexToStartFrom >= values.length || indexToStartFrom === undefined) ? 1 : indexToStartFrom; 

  for (var i = startIndex; i < values.length; i++) {
    var row = values[i];
    var isShared = row[map.compartido]; 

    if (isShared === false || String(isShared).toUpperCase() === 'FALSE' || isShared === '') {
      
      var fileId = row[map.idImagen]; 
      var imageUrl = fileId ? 'https://drive.google.com/uc?export=view&id=' + fileId : null;
      var rowIndex = i + 1; 
      
      var productData = {
        currentIndex: rowIndex, 
        totalItems: totalProducts,     
        rowIndex: rowIndex, 
        sku: row[map.sku],         
        marca: row[map.marca],       
        modelo: row[map.modelo],      
        genero: row[map.genero],      
        precio: row[map.precio],      
        imageLink: imageUrl
      };
      
      productData.formattedDescription = formatDescription(productData);
      return productData;
    }
  }
  
  if (startIndex > 1) {
    return getNextProduct(1); 
  }
  
  return { message: '¡Todos los productos han sido compartidos!' };
}


function skipProduct(rowIndex) {
  var nextProduct = getNextProduct(rowIndex + 1);
  if (!nextProduct.error && nextProduct.rowIndex) {
    saveUserProgress(nextProduct.rowIndex);
  }
  return nextProduct;
}


function markAsShared(rowIndex) {
  var sheet;
  var map;
  try {
    sheet = getSourceSheet();
    map = getHeaderMap(sheet); 
  } catch (e) {
    return { status: 'error', message: e.message };
  }
  
  sheet.getRange(rowIndex, map.compartido + 1).setValue(true);
  var nextProductData = getNextProduct(rowIndex + 1);
  
  if (!nextProductData.error && nextProductData.rowIndex) {
    saveUserProgress(nextProductData.rowIndex);
  }
  
  return { status: 'success', message: 'Producto marcado como compartido.', nextProduct: nextProductData };
}

/**
 * ⚡ ¡NUEVO!
 * Actualiza una o más celdas de un producto.
 * @param {number} rowIndex La fila a editar (1-based).
 * @param {Object} dataToUpdate Un objeto, ej: { "precio": 150000, "sku": "NUEVO-SKU" }
 */
function updateProduct(rowIndex, dataToUpdate) {
  try {
    const sheet = getSourceSheet();
    const map = getHeaderMap(sheet);

    // Mapeamos los campos que permitimos editar desde el mapa de cabeceras
    const allowedToEdit = {
      "sku": map.sku,
      "marca": map.marca,
      "modelo": map.modelo,
      "genero": map.genero,
      "precio": map.precio
    };

    for (const key in dataToUpdate) {
      const colIndex = allowedToEdit[key]; // colIndex es 0-based
      
      if (colIndex !== undefined) {
        // getRange es 1-based, por eso (colIndex + 1)
        sheet.getRange(rowIndex, colIndex + 1).setValue(dataToUpdate[key]);
      } else {
        Logger.log(`Intento de editar campo no permitido o no encontrado: ${key}`);
      }
    }
    
    return { status: "success", message: "Producto actualizado." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
