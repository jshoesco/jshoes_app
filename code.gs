/**
 * Se ejecuta automáticamente al abrir la hoja para crear el menú.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('FINANZAS RÁPIDO')
      .addItem('Abrir Formulario de Pedidos', 'mostrarFormularioPedidos')
      .addToUi();
}

/**
 * Muestra el formulario HTML de Pedidos como una ventana de diálogo modal.
 */
function mostrarFormularioPedidos() {
  var html = HtmlService.createHtmlOutputFromFile('Formulario')
      .setWidth(600) 
      .setHeight(650) 
      .setTitle('Registro de Pedidos/Rechazos');
      
  SpreadsheetApp.getUi().showModalDialog(html, 'Registro de Pedidos/Rechazos'); 
}

// =========================================================================
// FUNCIÓN ESTÁNDAR PARA DESPLIEGUE COMO APP WEB
// =========================================================================

/**
 * Función requerida para el despliegue como aplicación web.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Formulario')
      .evaluate()
      .setTitle('Sistema de Pedidos')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


// =========================================================================
// CONFIGURACIÓN CLAVE: IDs de los libros de datos (CRÍTICO)
// =========================================================================

/**
 * ID del documento de Google Sheet donde se encuentra tu lista de productos ACTIVOS (Fuente Primaria).
 */
const ID_LIBRO_PRODUCTOS = '1pbjnxZtYwnnvBWQbbzdLc2XCEO9qaXTyOjne8PNgz8I'; 

/**
 * ID del documento de Google Sheet donde se encuentran los SKUs pasados/históricos (Fuente Secundaria).
 */
const ID_LIBRO_SKUS_PASADOS = '1rGh8E7kOiXnm4NhA9VZLmDIA3m3HDfbRf-QvL1WVdek';

/**
 * ID del documento de Google Sheet donde se encuentran las tarifas de envío.
 */
const ID_LIBRO_TARIFAS = '1f-fDYQUPD8S8M-KevcvCJm57s93OI1aC9alINnHlnds'; 


// =========================================================================
// FUNCIÓN DE LECTURA DE PRODUCTOS - LÓGICA DE BÚSQUEDA DUAL 
// =========================================================================

/**
 * [HELPER] Lee la hoja especificada de un libro y devuelve una lista de objetos de producto.
 * Retorna un [] (array vacío) en caso de fallo, para permitir la concatenación segura.
 */
function _leerProductosDeLibro(idLibro, nombreHoja) {
    var productosDetalles = [];
    try {
        var libro = SpreadsheetApp.openById(idLibro); 
        var hoja = libro.getSheetByName(nombreHoja);

        if (!hoja) {
            Logger.log(`ERROR: La hoja '${nombreHoja}' no fue encontrada en el documento ID: ${idLibro}`);
            return []; 
        }
        
        var ultimaFila = hoja.getLastRow();
        if (ultimaFila <= 1) {
            Logger.log(`ERROR: La hoja '${nombreHoja}' no contiene datos (o solo el encabezado). ID: ${idLibro}`);
            return []; 
        }
        
        // Leemos 11 columnas a partir de la columna A (A2:K)
        var rangoDetalles = hoja.getRange(2, 1, ultimaFila - 1, 11).getValues();
        
        rangoDetalles.forEach(function(fila) {
          // Índices absolutos de la hoja: ID PRODUCTO (2), Proveedor(1), Marca(4), Modelo(5), Precio(10), Costo(8)
          
          var idProducto = String(fila[2]).trim(); // ID Producto Corto (el que usas para el autocomplete)
          var marca = String(fila[4]).trim() || 'Sin Marca'; 
          var modelo = String(fila[5]).trim() || 'Sin Modelo'; 
          var proveedor = String(fila[1]).replace(/\s/g, ' ').trim() || 'Sin Proveedor'; 
          var precio = parseFloat(fila[10]) || 0; // PRECIO en Columna K (índice 10)
          var costo = parseFloat(fila[8]) || 0; // COSTO BASE en Columna I (índice 8)
          
          if (idProducto !== "") {
            var etiquetaVisible = `${idProducto} - ${marca} - ${modelo} (${proveedor})`; 
            
            productosDetalles.push({
              sku: idProducto, // Usar idProducto como SKU para el front-end (ya que es el ID corto)
              etiqueta: etiquetaVisible,
              marca: marca,  
              modelo: modelo,
              proveedor: proveedor,
              precio: precio,
              costo: costo 
            });
          }
        });
        
        return productosDetalles;

    } catch (e) {
        Logger.log(`Error crítico al leer productos del libro ${idLibro}: ${e.toString()}`);
        return []; 
    }
}


/**
 * Lee y combina IDs de Producto disponibles del catálogo activo y del histórico.
 */
function getSkusDisponibles() { 
  const NOMBRE_HOJA_PRODUCTOS_ACTIVOS = 'productos';
  const NOMBRE_HOJA_PRODUCTOS_PASADOS = 'Hoja 1'; 

  // 1. Leer productos activos (Fuente Primaria)
  var idsActivos = _leerProductosDeLibro(ID_LIBRO_PRODUCTOS, NOMBRE_HOJA_PRODUCTOS_ACTIVOS);
  
  // 2. Leer IDs pasados (Fuente Secundaria)
  var idsPasados = _leerProductosDeLibro(ID_LIBRO_SKUS_PASADOS, NOMBRE_HOJA_PRODUCTOS_PASADOS);

  // 3. Combinar las listas
  var idsCombinados = idsActivos.concat(idsPasados);

  // 4. Comprobar si hay resultados
  if (idsCombinados.length > 0) {
    Logger.log(`INFO: ${idsActivos.length} IDs activos y ${idsPasados.length} IDs pasados cargados. Total: ${idsCombinados.length}`);
    return idsCombinados;
  }
  
  // 5. Fallo total
  var errorMsg = "ERROR: No se pudo obtener la lista de IDs de Producto en ningún catálogo o ambos están vacíos. Revisa los IDs de libros y los nombres de hojas.";
  Logger.log(errorMsg);
  return errorMsg; 
}


// =========================================================================
// FUNCIÓN DE LECTURA DE TARIFAS
// =========================================================================

/**
 * Lee la hoja 'inputs' del libro de tarifas y devuelve un mapa de Ciudad -> Tarifa.
 */
function getTarifasCiudades() {
  var tarifas = [];
  try {
    var libroTarifas = SpreadsheetApp.openById(ID_LIBRO_TARIFAS);
    var hojaTarifas = libroTarifas.getSheetByName('inputs'); 

    if (!hojaTarifas) {
      return "ERROR: La hoja 'inputs' no fue encontrada en el documento de tarifas.";
    }

    var ultimaFila = hojaTarifas.getLastRow();
    if (ultimaFila <= 1) {
      return "ERROR: La hoja 'inputs' no contiene datos (o solo el encabezado).";
    }

    // Leemos 2 columnas a partir de la columna A (A2:B): Ciudad (0), Tarifa (1)
    var rangoTarifas = hojaTarifas.getRange(2, 1, ultimaFila - 1, 2).getValues();

    rangoTarifas.forEach(function(fila) {
      var ciudad = String(fila[0]).trim(); // Columna A (índice 0)
      var tarifa = parseFloat(fila[1]) || 0; // Columna B (índice 1)

      if (ciudad !== "") {
        tarifas.push({
          ciudad: ciudad,
          tarifa: tarifa 
        });
      }
    });

    return tarifas;

  } catch (e) {
    if (e.message.includes('No se encuentra la hoja de cálculo con el ID')) {
      return "ERROR: Verifica que el ID del libro de tarifas ('" + ID_LIBRO_TARIFAS + "') sea correcto y que tengas permisos de acceso.";
    }
    return "ERROR: Hubo un problema al leer la hoja de tarifas: " + e.toString();
  }
}

// =========================================================================
// FUNCIÓN CRÍTICA PARA EL AUTOCOMPLETE DE RECHAZO/DEVOLUCIÓN
// =========================================================================

/**
 * Obtiene la lista de ID PRODUCTO ÚNICO de la Columna C de la hoja 'Pedidos'.
 * Usado para la Pestaña "Rechazo/Devolución"
 */
function getProductosRechazoDisponibles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pedidos'); 
    
    if (!sheet) {
      return "ERROR: Hoja 'Pedidos' no encontrada. Verifica el nombre.";
    }

    // Columna C (ID Producto Único)
    const idCol = 3; 
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return []; 
    }

    // Rango: Columna C desde Fila 2 hasta la última fila
    const dataRange = sheet.getRange(2, idCol, lastRow - 1, 1); 
    const values = dataRange.getValues();
    
    // Aplanar el array y filtrar celdas vacías
    const idProductosUnicos = values.map(row => String(row[0]).trim()).filter(id => id !== "");
    
    return idProductosUnicos;

  } catch (e) {
    Logger.log("Error en getProductosRechazoDisponibles: " + e.message);
    return "ERROR: Falló la ejecución en el servidor. Mensaje: " + e.message;
  }
}

// =========================================================================
// FUNCIONES CRUD DE EDICIÓN (COMPLETADO)
// =========================================================================

/**
 * Obtiene la lista de ID Pedido (Columna B) únicos de la hoja 'Pedidos'.
 * Usado para el campo de edición/búsqueda.
 */
function getIdsPedidosDisponibles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pedidos'); 
    
    if (!sheet) {
      return []; 
    }

    // Columna B (ID Pedido)
    const idCol = 2; 
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return []; 
    }

    // Rango: Columna B desde Fila 2 hasta la última fila
    const dataRange = sheet.getRange(2, idCol, lastRow - 1, 1); 
    const values = dataRange.getValues();
    
    // 1. Aplanar el array, convertir a string y limpiar espacios.
    const allIds = values.map(row => String(row[0]).replace(/\s/g, '').trim().toUpperCase()).filter(id => id !== "");
    
    // 2. Obtener valores únicos.
    const uniqueIds = [...new Set(allIds)];
    
    return uniqueIds;

  } catch (e) {
    Logger.log("Error en getIdsPedidosDisponibles: " + e.message);
    return [];
  }
}

/**
 * Busca y retorna los detalles de TODAS las líneas de pedido que coinciden con el ID Pedido.
 * Esto es necesario porque un ID Pedido (Columna B) puede tener múltiples líneas.
 * @param {string} idPedido El ID de pedido completo (Columna B).
 * @returns {Array<Object>|null} Un array de objetos (cada uno es una línea de pedido) o null si no se encuentra.
 */
function getPedidoParaEdicion(idPedido) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pedidos');

    if (!sheet) {
      return { error: "Hoja 'Pedidos' no encontrada. Verifica el nombre." };
    }

    const idBuscado = String(idPedido).replace(/\s/g, '').trim().toUpperCase();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return null;
    }

    // Leemos la data completa (asumimos 18 columnas: A a R)
    const data = sheet.getRange(2, 1, lastRow - 1, 18).getValues(); 
    
    // Encabezados para crear objetos con clave-valor (ajustar si las columnas cambian)
    const headers = [
      'Fecha Pedido', 'ID Pedido', 'ID Producto Único', 'Cliente', 'ID_PRODUCTO', 
      'Contacto', 'Ciudad Cliente', 'Ciudad Entrega', 'Talla', 'Marca-Modelo', 
      'Proveedor', 'Cantidad', 'Precio Unitario', 'Precio Total', 'Estado', 'Cantidad Rechazada',
      'Número Guía', 'Fecha Envío' // Columnas Q y R
    ];
    
    const pedidosEncontrados = [];
    
    for (let i = 0; i < data.length; i++) {
      const fila = data[i];
      const idEnHoja = String(fila[1]).replace(/\s/g, '').trim().toUpperCase(); // Columna B: ID Pedido

      if (idEnHoja === idBuscado) {
        const rowObject = {};
        rowObject.rowIndex = i + 2; // Fila real en la hoja
        
        headers.forEach((header, index) => {
            rowObject[header] = fila[index];
        });
        
        // Convertir la fecha a formato ISO string para el front-end (YYYY-MM-DD)
        if (rowObject['Fecha Pedido'] instanceof Date) {
            rowObject['Fecha Pedido'] = Utilities.formatDate(rowObject['Fecha Pedido'], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
        }
        // Convertir Fecha Envío también si existe
        if (rowObject['Fecha Envío'] instanceof Date) {
            rowObject['Fecha Envío'] = Utilities.formatDate(rowObject['Fecha Envío'], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
        }

        pedidosEncontrados.push(rowObject);
      }
    }

    if (pedidosEncontrados.length > 0) {
        return pedidosEncontrados;
    }
    
    return null; // No se encontró el pedido
    
  } catch (e) {
    Logger.log("Error en getPedidoParaEdicion: " + e.message);
    return { error: "Error de servidor al buscar el pedido: " + e.message };
  }
}

/**
 * Función que actualiza las filas de pedido modificadas.
 * IMPLEMENTACIÓN FINAL.
 * @param {Array<Object>} datosActualizados Array de objetos con datos y rowIndex a actualizar.
 */
function actualizarPedido(datosActualizados) {
    if (!datosActualizados || datosActualizados.length === 0) {
        return { success: false, message: "No se recibieron datos para actualizar." };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pedidos');

    if (!sheet) {
        return { success: false, message: "Hoja 'Pedidos' no encontrada." };
    }

    try {
        const numCols = 18; // A hasta R

        datosActualizados.forEach(function(data) {
            const rowIndex = data.rowIndex;
            if (!rowIndex || rowIndex < 2) {
                Logger.log("Advertencia: rowIndex inválido/omitido. Fila no actualizada.");
                return;
            }
            
            // Los datos vienen del front-end en el orden correcto
            // Recalculamos Precio Total (Columna N) para asegurar consistencia
            const precioTotalLinea = (parseFloat(data['Cantidad']) || 0) * (parseFloat(data['Precio Unitario']) || 0);
            
            // Convertir el contacto a formato de texto para evitar problemas con el '+'
            const contactoTexto = "'" + (data['Contacto'] || ''); 

            // Convertir Fecha Envío si viene
            let fechaEnvioValor = data['Fecha Envío'] ? new Date(data['Fecha Envío']) : '';

            const values = [
                // Aseguramos que la fecha se convierta a objeto Date si viene como string
                new Date(data['Fecha Pedido']), // A: Fecha Pedido
                data['ID Pedido'],             // B: ID Pedido
                data['ID Producto Único'],     // C: ID Producto Único
                data['Cliente'],               // D: Cliente
                data['ID_PRODUCTO'],           // E: ID_PRODUCTO
                contactoTexto,                 // F: Contacto
                data['Ciudad Cliente'],        // G: Ciudad Cliente
                data['Ciudad Entrega'],        // H: Ciudad Entrega
                data['Talla'],                 // I: Talla
                data['Marca-Modelo'],          // J: Marca-Modelo (NO editable)
                data['Proveedor'],             // K: Proveedor (NO editable)
                parseFloat(data['Cantidad']) || 0,        // L: Cantidad
                parseFloat(data['Precio Unitario']) || 0, // M: Precio Unitario
                precioTotalLinea,              // N: Precio Total (RECALCULADO)
                data['Estado'],                // O: Estado
                parseFloat(data['Cantidad Rechazada']) || 0, // P: Cantidad Rechazada
                data['Número Guía'] || '',     // Q: Número Guía
                fechaEnvioValor                // R: Fecha Envío
            ];
            
            // Escribir la fila completa
            sheet.getRange(rowIndex, 1, 1, numCols).setValues([values]);
        });
        
        return { success: true, message: `Pedido ${datosActualizados[0]['ID Pedido']} actualizado con éxito. (${datosActualizados.length} línea(s) modificada(s))` };

    } catch (e) {
        Logger.log(`Error crítico al actualizar el pedido: ${e.toString()}`);
        return { success: false, message: "Error de servidor al actualizar el pedido: " + e.message };
    }
}

// =========================================================================
// FUNCIÓN PRINCIPAL: Guardar Pedido 
// =========================================================================

/**
 * Función que agrupa los productos por Proveedor, asigna IDs ÚNICOS y CONSECUTIVOS
 * basados en la primera letra del ID Producto, y usa escritura por lotes.
 */
function guardarPedido(datos) {
  
  if (!datos || !datos.productos || !Array.isArray(datos.productos) || datos.productos.length === 0) {
    throw new Error("No se pudo registrar el pedido: La estructura de datos es inválida o no contiene productos.");
  }
  
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  
  // *** VALOR FIJO ***
  const ESTADO_FIJO = 'Pendiente'; 

  // Crear Hoja y Encabezados si es necesario (18 columnas)
  if (!hoja || hoja.getLastRow() === 0) {
    if (!hoja) {
      hoja = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Pedidos');
    }
    hoja.clear();
    // ENCABEZADO: AÑADIMOS Col Q (17) y Col R (18)
    hoja.appendRow(['Fecha Pedido', 'ID Pedido', 'ID Producto Único', 'Cliente', 'ID PRODUCTO', 'Contacto', 'Ciudad Cliente', 'Ciudad Entrega', 'Talla', 'Marca-Modelo', 'Proveedor', 'Cantidad', 'Precio Unitario', 'Precio Total', 'Estado', 'Cantidad Rechazada', 'Número Guía', 'Fecha Envío']);
    hoja.setFrozenRows(1); // Congelar encabezados
  }
  
  // 1. Agrupar productos por proveedor
  var productosPorProveedor = {};
  datos.productos.forEach(function(producto) {
    var proveedor = producto.proveedor.replace(/\s/g, ' ').trim() || 'SIN_PROVEVEDOR';
    if (!productosPorProveedor[proveedor]) {
      productosPorProveedor[proveedor] = [];
    }
    productosPorProveedor[proveedor].push(producto);
  });

  // 2. Calcular la secuencia base y prefijo de fecha
  var fechaPedido = new Date(datos.fechaPedido + 'T00:00:00');
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  // Formato YYMMDD
  var fechaFormato = Utilities.formatDate(fechaPedido, timeZone, 'yyMMdd');
  
  // ----------------------------------------------------------------------
  // 💡 LÓGICA AJUSTADA: Encontrar la secuencia máxima global (sin importar prefijo)
  // ----------------------------------------------------------------------
  var maxGlobalSecuencia = 0;
  var ultimaFila = hoja.getLastRow();

  if (ultimaFila >= 2) {
    var rangoIds = hoja.getRange(2, 2, ultimaFila - 1, 1).getValues();
    rangoIds.forEach(function(fila) {
      var idCompleto = String(fila[0]);
      var partesId = idCompleto.split('-'); // Se espera un formato como YYMMDD-X-001
      
      // Debe tener 3 partes (Fecha, Prefijo, Secuencia)
      if (partesId.length === 3) {
        var secuencia = parseInt(partesId[2]); // Secuencia es el final: 001, 007, etc.
        
        // Almacenar el número más alto encontrado, sin importar la fecha ni el prefijo
        if (!isNaN(secuencia) && secuencia > maxGlobalSecuencia) {
          maxGlobalSecuencia = secuencia;
        }
      }
    });
  }

  // Se inicia la secuencia actual con la máxima encontrada. Se incrementará en el bucle.
  var currentGlobalSecuencia = maxGlobalSecuencia;

  // 3. Construir la matriz de datos final (rowsToWrite)
  var rowsToWrite = [];
  var proveedoresUnicos = Object.keys(productosPorProveedor);
  
  proveedoresUnicos.forEach(function(proveedor) {
    var productosDelProveedor = productosPorProveedor[proveedor];
    
    // Usar la primera letra del ID Producto Corto (sku) para el prefijo del ID de Pedido.
    var primerSku = productosDelProveedor[0].sku;
    var prefijoProveedor = primerSku.charAt(0).toUpperCase(); 
    
    // INCREMENTO GLOBAL: La secuencia aumenta por cada LOTE/ID de pedido único que se va a crear
    currentGlobalSecuencia++;
    
    // Generar el nuevo ID de Pedido con la secuencia global siguiente
    var idPedidoBase = `${fechaFormato}-${prefijoProveedor}-${String(currentGlobalSecuencia).padStart(3, '0')}`;
    
    productosDelProveedor.forEach(function(producto) {
      // Calcular costos y precios
      var precioTotalLinea = producto.cantidad * producto.precioUnitario;
      var idProductoUnico = `${idPedidoBase}-${producto.sku}-${generateRandomCode(4)}-${producto.cantidad}`;

      // ------------------------------------
      // FORMATO DE DATOS DE TEXTO
      // ------------------------------------
      const clienteFormateado = toProperCase(producto.cliente);
      const ciudadClienteFormateada = toProperCase(producto.ciudadCliente);
      const tallaFormateada = toProperCase(producto.talla);
      const marcaModelo = `${producto.marca} - ${producto.modelo}`;
      
      // Mantiene el formato de texto para el contacto (solución del '+' anterior)
      const contactoTexto = "'" + (producto.contacto || ''); 
      // ------------------------------------
      
      var nuevaFila = [
        fechaPedido, // A: Fecha Pedido
        idPedidoBase, // B: ID Pedido
        idProductoUnico, // C: ID Producto Único (Clave Única)
        clienteFormateado || '', // D: Cliente
        producto.sku, // E: ID Producto
        contactoTexto, // F: Contacto
        ciudadClienteFormateada || '', // G: Ciudad Cliente
        producto.ciudadEntrega || '', // H: Ciudad Entrega
        tallaFormateada || '', // I: Talla
        marcaModelo || '', // J: Marca-Modelo
        producto.proveedor || '', // K: Proveedor
        producto.cantidad || 0, // L: Cantidad
        producto.precioUnitario || 0, // M: Precio Unitario
        precioTotalLinea || 0, // N: Precio Total
        ESTADO_FIJO, // O: Estado
        0, // P: Cantidad Rechazada (0 por defecto)
        '', // Q: Número Guía (Vacío por defecto)
        ''  // R: Fecha Envío (Vacío por defecto)
      ];
      rowsToWrite.push(nuevaFila);
    });
  });

  // 4. Escritura por lotes (eficiente)
  if (rowsToWrite.length > 0) {
    hoja.getRange(hoja.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
    return { success: true, idPedido: rowsToWrite[0][1] };
  } else {
    throw new Error("No se generó ninguna fila para escribir.");
  }
}

// =========================================================================
// FUNCIONES DE UTILIDAD Y CÁLCULO
// =========================================================================

/**
 * Genera un código alfanumérico aleatorio.
 * @param {number} length Longitud del código.
 * @returns {string} Código aleatorio.
 */
function generateRandomCode(length) {
  var characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var result = '';
  for (var i = 0; i < length; i++) {
    result += characters.charAt(Math.floor(Math.random() * characters.length));
  }
  return result;
}

/**
 * Convierte una cadena a formato Nombre Propio (Title Case) y elimina espacios extra.
 * Reemplaza el uso de .trim() en los campos especificados.
 * @param {string} str La cadena de entrada.
 * @returns {string} La cadena formateada.
 */
function toProperCase(str) {
  if (!str) return '';
  // 1. Eliminar espacios innecesarios (al inicio/final y múltiples internos)
  str = String(str).replace(/\s+/g, ' ').trim();
  // 2. Aplicar mayúsculas a la primera letra de cada palabra
  return str.toLowerCase().split(' ').map(function(word) {
    return word.charAt(0).toUpperCase() + word.slice(1);
  }).join(' ');
}

// --------------------------------------------------------------------------
// Funciones relacionadas con Rechazo/Devolución (CORREGIDAS)
// --------------------------------------------------------------------------

/**
 * Busca y retorna la línea de pedido que coincide con el ID Producto Único (Columna C).
 * Usado para verificar la disponibilidad de rechazo/devolución.
 * *** MODIFICADO: Ahora devuelve marcaModelo ***
 */
function getProductoParaRechazo(idProductoUnico) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pedidos');

    if (!sheet) {
      return { error: "Hoja 'Pedidos' no encontrada." };
    }

    const idBuscado = String(idProductoUnico).replace(/\s/g, '').trim().toUpperCase();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return null;
    }
    
    // Leemos solo las columnas necesarias para no sobrecargar: Columna C (ID Único), L (Cantidad), P (Cant. Rechazada)
    const data = sheet.getRange(2, 3, lastRow - 1, 14).getValues(); // Leemos desde C hasta P

    for (let i = 0; i < data.length; i++) {
      const fila = data[i];
      const idEnHoja = String(fila[0]).replace(/\s/g, '').trim().toUpperCase(); // Columna C (índice 0 en el sub-rango)

      if (idEnHoja === idBuscado) {
        const cantidadOriginal = parseInt(fila[9]) || 0; // Columna L (índice 9 en el sub-rango)
        const cantidadRechazada = parseInt(fila[13]) || 0; // Columna P (índice 13 en el sub-rango)
        
        return {
          rowIndex: i + 2,
          cantidadTotal: cantidadOriginal,
          cantidadYaRechazada: cantidadRechazada,
          unidadesDisponibles: cantidadOriginal - cantidadRechazada,
          marcaModelo: String(fila[7]) || '' // Col J (Marca-Modelo) es el índice 7
        };
      }
    }
    
    return null; // No se encontró el producto
    
  } catch (e) {
    Logger.log("Error en getProductoParaRechazo: " + e.message);
    return { error: "Error de servidor al buscar el producto: " + e.message };
  }
}

/**
 * Actualiza la Columna P (Cantidad Rechazada) de la línea de pedido.
 * También registra el artículo en 'Inventario Devolucion'.
 */
function registrarRechazo(idProductoUnico, cantidad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pedidos');

  const producto = getProductoParaRechazo(idProductoUnico);

  if (!producto || producto.error || producto.unidadesDisponibles < cantidad) {
    throw new Error("Validación fallida: El producto no existe o la cantidad a rechazar supera las unidades disponibles.");
  }
  
  const nuevaCantidadRechazada = producto.cantidadYaRechazada + cantidad;
  
  // 1. Actualizar la hoja Pedidos (Columna P)
  const rangoP = sheet.getRange(producto.rowIndex, 16); // Columna P
  rangoP.setValue(nuevaCantidadRechazada);

  // 2. Registrar en la hoja de inventario (Necesitamos los detalles del producto para esto)
  // Leemos hasta la Columna P (16)
  const filaCompleta = sheet.getRange(producto.rowIndex, 1, 1, 16).getValues()[0];
  
  // Objeto de datos (Data Object) para evitar errores de orden
  const datosRechazo = {
    idUnico: idProductoUnico,
    sku: filaCompleta[4],         // Col E (ID PRODUCTO)
    marcaModelo: filaCompleta[9], // Col J (Marca-Modelo)
    talla: filaCompleta[8],       // Col I (Talla)
    cantidad: cantidad,           // Del formulario
    proveedor: filaCompleta[10],  // Col K (Proveedor)
    cliente: filaCompleta[3]      // Col D (Cliente)
  };

  // Usamos la nueva función helper para registro.
  _registrarEnInventarioDevolucion(datosRechazo);
  
  return { success: true, newTotalRechazo: nuevaCantidadRechazada };
}


/**
 * [HELPER] Registra un artículo rechazado en la hoja 'Inventario Devolucion'.
 * (CORREGIDA para tu estructura de 9 columnas - A hasta I)
 */
function _registrarEnInventarioDevolucion(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Inventario Devolucion');

  // Crear la hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet('Inventario Devolucion');
    // ENCABEZADOS DE 9 COLUMNAS (A-I)
    sheet.appendRow(['ID Producto Único', 'Sku', 'Marca-Modelo', 'Talla', 'Cantidad Disponible', 'Proveedor', 'Cliente Original', 'Status', 'Fecha Ingreso']);
    sheet.setFrozenRows(1);
  }

  const filaExistente = _buscarFilaEnInventarioDevolucion(datos.idUnico);

  if (filaExistente) {
    // Si ya existe, actualiza la cantidad (Columna E) y el status (Col H)
    const nuevaCantidad = filaExistente.cantidadActual + datos.cantidad;
    sheet.getRange(filaExistente.rowIndex, 5).setValue(nuevaCantidad); // Col E (Cantidad)
    sheet.getRange(filaExistente.rowIndex, 8).setValue('Disponible'); // Col H (Status)
  } else {
    // Si no existe, crea una nueva fila (9 columnas)
    // El orden en este array debe coincidir EXACTAMENTE con tus 9 encabezados
    const nuevaFila = [
      datos.idUnico,   // A: ID Producto Único
      datos.sku,       // B: Sku
      datos.marcaModelo, // C: Marca-Modelo
      datos.talla,     // D: Talla
      datos.cantidad,  // E: Cantidad Disponible
      datos.proveedor, // F: Proveedor
      datos.cliente,   // G: Cliente Original
      'Disponible',    // H: Status
      new Date()       // I: Fecha Ingreso
    ];
    sheet.appendRow(nuevaFila);
  }
}

/**
 * [HELPER] Busca un ID Producto Único en 'Inventario Devolucion'.
 * (CORREGIDA para tu estructura de 9 columnas - A hasta I)
 * *** MODIFICADO: Ahora devuelve marcaModelo ***
 */
function _buscarFilaEnInventarioDevolucion(idProductoUnico) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Inventario Devolucion');

  if (!sheet || sheet.getLastRow() < 2) return null;

  const lastRow = sheet.getLastRow();
  // Leemos Columna A (ID Único), E (Cantidad), H (Status)
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues(); // Leemos A:H (8 columnas de datos)

  for (let i = 0; i < data.length; i++) {
    const fila = data[i];
    const idEnHoja = String(fila[0]).replace(/\s/g, '').trim().toUpperCase(); // Columna A
    const idBuscado = String(idProductoUnico).replace(/\s/g, '').trim().toUpperCase();

    if (idEnHoja === idBuscado) {
      return {
        rowIndex: i + 2,
        marcaModelo: String(fila[2]).trim(),   // Col C (índice 2)
        cantidadActual: parseInt(fila[4]) || 0, // Col E (índice 4)
        statusActual: String(fila[7]).trim()   // Col H (índice 7)
      };
    }
  }
  return null;
}

// --------------------------------------------------------------------------
// Funciones relacionadas con Venta Devolución (CORREGIDAS)
// --------------------------------------------------------------------------

/**
 * Busca y retorna la cantidad disponible en 'Inventario Devolucion'.
 * (CORREGIDA para tu estructura de 9 columnas)
 * *** MODIFICADO: Ahora devuelve marcaModelo ***
 */
function getDisponibilidadVentaDevolucion(idProductoUnico) {
  const fila = _buscarFilaEnInventarioDevolucion(idProductoUnico);
  
  if (fila) {
    return {
      disponible: fila.cantidadActual,
      rowIndex: fila.rowIndex,
      status: fila.statusActual,
      marcaModelo: fila.marcaModelo
    };
  }
  return { disponible: 0, rowIndex: -1, status: '', marcaModelo: '' };
}

/**
 * Registra la venta de un artículo del inventario de devoluciones.
 * (CORREGIDA para tu estructura de 9 columnas)
 */
function registrarVentaDevolucion(datosVenta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetInventario = ss.getSheetByName('Inventario Devolucion');
  let sheetLog = ss.getSheetByName('Ventas Devolucion Log');

  // 1. Validar y obtener disponibilidad
  const producto = getDisponibilidadVentaDevolucion(datosVenta.idProductoVenta);
  const cantidadVendida = parseInt(datosVenta.cantidadVenta);

  if (producto.disponible < cantidadVendida) {
    throw new Error(`Cantidad insuficiente. Solo hay ${producto.disponible} unidades disponibles para el ID: ${datosVenta.idProductoVenta}`);
  }
  
  if (producto.statusActual !== 'Disponible') {
     throw new Error(`Este producto no está 'Disponible' (Estado actual: ${producto.statusActual}).`);
  }

  // 2. Actualizar Inventario (restar cantidad y actualizar status)
  const nuevaCantidad = producto.disponible - cantidadVendida;
  sheetInventario.getRange(producto.rowIndex, 5).setValue(nuevaCantidad); // Col E (Cantidad)

  let nuevoEstado = 'Disponible';
  if (nuevaCantidad <= 0) {
    nuevoEstado = 'Vendido';
    sheetInventario.getRange(producto.rowIndex, 8).setValue(nuevoEstado); // Col H (Status)
  }

  // 3. Registrar en Log (Crear la hoja si no existe)
  if (!sheetLog || sheetLog.getLastRow() === 0) {
    if (!sheetLog) {
      sheetLog = ss.insertSheet('Ventas Devolucion Log');
    }
    sheetLog.clear();
    sheetLog.appendRow(['Fecha Venta', 'Cliente', 'Contacto', 'Dirección Entrega', 'Ciudad Destino', 'ID Producto Único', 'Cantidad', 'Precio Unitario']);
    sheetLog.setFrozenRows(1);
  }

  const fechaVenta = new Date();
  
  // ------------------------------------
  // FORMATO DE DATOS DE TEXTO
  // ------------------------------------
  const clienteFormateado = toProperCase(datosVenta.cliente);
  const direccionFormateada = toProperCase(datosVenta.direccionEntrega);
  const ciudadFormateada = toProperCase(datosVenta.ciudadEntrega);
  // ------------------------------------

  // Mantiene el formato de texto para el contacto (solución del '+' anterior)
  const contactoTexto = "'" + (datosVenta.contacto || ''); 
  
  const nuevaFila = [
      fechaVenta,
      clienteFormateado || '', // B (Cliente)
      contactoTexto,            // C (Contacto)
      direccionFormateada || '', // D (Dirección Entrega)
      ciudadFormateada || '',    // E (Ciudad Destino)
      datosVenta.idProductoVenta, // ID CRÍTICO
      parseInt(datosVenta.cantidadVenta) || 0,
      parseFloat(datosVenta.precioUnitario) || 0
  ];
  
  sheetLog.appendRow(nuevaFila);

  return { success: true, nuevaCantidadDisponible: nuevaCantidad, newStatus: nuevoEstado };
}


// =========================================================================
// FUNCIÓN PARA REGISTRAR GUÍA DE ENVÍO
// =========================================================================

/**
 * Registra el número de guía y la fecha de envío para TODAS las líneas de un ID Pedido.
 * Escribe en las columnas O (Estado), Q (Guía) y R (Fecha).
 * @param {string} idPedido El ID de pedido (Columna B).
 * @param {string} numeroGuia El número de guía a registrar.
 * @param {string} fechaEnvio La fecha de envío (en formato YYYY-MM-DD).
 * @returns {object} Objeto con el resultado de la operación.
 */
function registrarGuiaEnvio(idPedido, numeroGuia, fechaEnvio) {
  if (!idPedido || !numeroGuia || !fechaEnvio) {
    return { success: false, message: "Faltan datos (ID Pedido, Guía o Fecha)." };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pedidos');

    if (!sheet) {
      return { success: false, message: "Hoja 'Pedidos' no encontrada." };
    }

    const idBuscado = String(idPedido).replace(/\s/g, '').trim().toUpperCase();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return { success: false, message: "La hoja 'Pedidos' está vacía." };
    }

    // Leemos solo la columna B (ID Pedido) para encontrar las filas
    const idRange = sheet.getRange(2, 2, lastRow - 1, 1);
    const idValues = idRange.getValues();
    
    let filasActualizadas = 0;
    const fechaEnvioDate = new Date(fechaEnvio + 'T00:00:00'); // Asegurar formato de fecha
    const ESTADO_NUEVO = 'Enviado'; // Nuevo estado

    // Iteramos para encontrar todas las coincidencias
    for (let i = 0; i < idValues.length; i++) {
      const idEnHoja = String(idValues[i][0]).replace(/\s/g, '').trim().toUpperCase();
      
      if (idEnHoja === idBuscado) {
        const rowIndex = i + 2; // Fila real en la hoja (offset +2)
        
        // *** CAMBIO: Escribir en Columna O (15) - ESTADO ***
        sheet.getRange(rowIndex, 15).setValue(ESTADO_NUEVO);
        
        // Escribir en Columna Q (17)
        sheet.getRange(rowIndex, 17).setValue(numeroGuia);
        
        // Escribir en Columna R (18)
        sheet.getRange(rowIndex, 18).setValue(fechaEnvioDate);
        
        filasActualizadas++;
      }
    }

    if (filasActualizadas > 0) {
      return { success: true, message: `Guía registrada con éxito para ${idPedido}. Estado actualizado a '${ESTADO_NUEVO}'. Se actualizaron ${filasActualizadas} línea(s).` };
    } else {
      return { success: false, message: `No se encontró el ID Pedido: ${idPedido}.` };
    }

  } catch (e) {
    Logger.log("Error en registrarGuiaEnvio: " + e.message);
    return { success: false, message: "Error de servidor: " + e.message };
  }
}

// =========================================================================
// FUNCIÓN PARA CORREGIR RECHAZO (NUEVA PESTAÑA 6 - CORREGIDA)
// =========================================================================

/**
 * Corrige un rechazo mal registrado.
 * Resta la cantidad del Inventario Devolucion (Col E).
 * Si llega a 0, elimina la fila del inventario.
 * Resta la cantidad de Pedidos (Col P).
 * Restaura el Status en Pedidos (Col O) a 'Pendiente' o 'Enviado'.
 */
function corregirRechazo(idProductoUnico, cantidadACorregir) {
  if (!idProductoUnico || !cantidadACorregir || cantidadACorregir <= 0) {
    return { success: false, message: "Datos inválidos. Se requiere ID y una cantidad mayor a 0." };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetInventario = ss.getSheetByName('Inventario Devolucion');
    const sheetPedidos = ss.getSheetByName('Pedidos');

    if (!sheetInventario || !sheetPedidos) {
      return { success: false, message: "No se encontraron las hojas 'Pedidos' o 'Inventario Devolucion'." };
    }
    
    cantidadACorregir = parseInt(cantidadACorregir);
    let mensajeResultado = '';

    // --- 1. Validar y Actualizar Inventario Devolucion ---
    const productoInv = _buscarFilaEnInventarioDevolucion(idProductoUnico);
    
    if (!productoInv) {
      return { success: false, message: "Error: El ID Producto Único no fue encontrado en 'Inventario Devolucion'." };
    }
    
    if (productoInv.cantidadActual < cantidadACorregir) {
      return { success: false, message: `Error: No se puede corregir ${cantidadACorregir} unidades. Solo hay ${productoInv.cantidadActual} disponibles en el inventario de devolución.` };
    }

    const nuevaCantidadInv = productoInv.cantidadActual - cantidadACorregir;
    
    // *** LÓGICA DE ELIMINACIÓN/ACTUALIZACIÓN CORREGIDA ***
    if (nuevaCantidadInv <= 0) {
      // Si la cantidad llega a 0, ELIMINA la fila
      sheetInventario.deleteRow(productoInv.rowIndex);
      mensajeResultado = `Inventario actualizado. La fila del producto fue eliminada (stock 0).`;
    } else {
      // Si aún queda stock, solo actualiza la cantidad en Col E
      sheetInventario.getRange(productoInv.rowIndex, 5).setValue(nuevaCantidadInv); // Col E (Cantidad)
      mensajeResultado = `Inventario actualizado. Stock restante: ${nuevaCantidadInv}.`;
    }
    // *** FIN DE LA LÓGICA DE ELIMINACIÓN ***


    // --- 2. Validar y Actualizar Hoja Pedidos ---
    const productoPed = getProductoParaRechazo(idProductoUnico); 
    
    if (!productoPed || productoPed.error) {
       Logger.log(`Error de Sincronización: ${idProductoUnico} existe en Inventario pero no se encontró en Pedidos.`);
       return { success: false, message: "Error de Sincronización: No se pudo encontrar el producto en la hoja 'Pedidos'." };
    }

    const filaPedido = productoPed.rowIndex;
    // Leer Col O (Status, 15), P (Cant. Rechazada, 16), Q (Guía, 17)
    const rangoPedido = sheetPedidos.getRange(filaPedido, 15, 1, 3); 
    const valoresPedido = rangoPedido.getValues()[0];
    
    const statusActualPedido = valoresPedido[0]; // Col O
    const cantRechazadaActual = parseInt(valoresPedido[1]) || 0; // Col P
    const guiaEnvio = valoresPedido[2]; // Col Q
    
    const nuevaCantRechazada = cantRechazadaActual - cantidadACorregir;
    
    if (nuevaCantRechazada <= 0) {
      // Si la corrección pone los rechazados en 0 o menos...
      sheetPedidos.getRange(filaPedido, 16).setValue(0); // Set Col P to 0
      
      // ...revertir el Status (Col O)
      if (guiaEnvio === '') {
           sheetPedidos.getRange(filaPedido, 15).setValue('Pendiente'); // Col O
           mensajeResultado += "\nEstado del Pedido: Revertido a 'Pendiente'.";
      } else {
           sheetPedidos.getRange(filaPedido, 15).setValue('Enviado'); // Col O
           mensajeResultado += "\nEstado del Pedido: Revertido a 'Enviado' (tenía guía).";
      }
      
    } else {
      // Si todavía quedan items rechazados (e.g., 3 rechazados, corrigió 1)
      sheetPedidos.getRange(filaPedido, 16).setValue(nuevaCantRechazada); // Set Col P to 2
      mensajeResultado += `\nPedido actualizado. ${nuevaCantRechazada} unidades aún marcadas como rechazadas.`;
    }

    return { success: true, message: mensajeResultado };

  } catch (e) {
    Logger.log("Error en corregirRechazo: " + e.message);
    return { success: false, message: "Error de servidor: " + e.message };
  }
}

// =========================================================================
// FUNCIÓN PARA OBTENER IDs DEL INVENTARIO (NUEVA)
// =========================================================================

/**
 * Obtiene la lista de ID PRODUCTO ÚNICO de la Columna A de la hoja 'Inventario Devolucion'.
 * Usado para las pestañas "Venta Devolución" y "Corregir Rechazo"
 */
function getIdsInventarioDisponibles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Inventario Devolucion'); 
    
    if (!sheet) {
      return []; // No es un error, puede que la hoja no exista aún
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return []; // Hoja vacía
    }

    // Rango: Columna A desde Fila 2 hasta la última fila
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 1); 
    const values = dataRange.getValues();
    
    // Aplanar el array y filtrar celdas vacías
    const idProductos = values.map(row => String(row[0]).trim()).filter(id => id !== "");
    
    return idProductos;

  } catch (e) {
    Logger.log("Error en getIdsInventarioDisponibles: " + e.message);
    return "ERROR: Falló la ejecución en el servidor. Mensaje: " + e.message;
  }
}
