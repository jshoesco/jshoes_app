// =============================================================================
// 1. CONFIGURACIÓN Y SEGURIDAD
// =============================================================================

// ✅ IDs y Email (Ya configurados)
const FINANZAS_SS_ID = "11ZBXJTMqxtMpd92w5IntNP4sXi0MEq_0TifeZ8iZ7XU";
const PEDIDOS_SS_ID = "1CbjkBbPKtzMR4yKowrCdhu-TpvSXBvi1lIVaKpAgBr0";
const SKUS_PASADOS_SS_ID = "1rGh8E7kOiXnm4NhA9VZLmDIA3m3HDfbRf-QvL1WVdek"; 
const ADMIN_EMAIL = "tennisandariegos@gmail.com"; 

// ✅ Configuración de la hoja de Pedidos (Confirmada)
const HOJA_DE_PEDIDOS = "Pedidos";
const HOJA_SKUS_PASADOS = "Hoja 1"; 
const COLUMNA_ID_PEDIDO = "ID Pedido";
const COLUMNA_ID_PRODUCTO = "ID Producto"; 
const COLUMNA_CLIENTE = "Cliente";
const COLUMNA_SKU = "SKU";
const COLUMNA_MARCA_MODELO = "Marca-Modelo";
const COLUMNA_PRECIO = "Precio Unitario"; 
const COLUMNA_ESTADO = "Estado"; 
const COLUMNA_COSTO = "Costo"; 

// ⚡ Índices del Libro de Respaldo (SkusPasados: Columna C=2, Columna I=8)
const RES_SKU_COL = 2; 
const RES_COSTO_COL = 8; 

// =============================================================================
// 2. SETUP DE INTERFAZ Y HOJA
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Finanzas")
    .addItem("Registrar Movimiento", "abrirForm")
    .addItem("Editar Pagos", "abrirForm") 
    .addSeparator()
    .addItem("⚡ [TEMPORAL] Configurar Hoja 'Detalle'", "_setupFinanzasSheet")
    .addToUi();
}

function _setupFinanzasSheet() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.openById(FINANZAS_SS_ID);
    const sheetName = "Movimiento_Detalle";
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      ui.alert(`Hoja "${sheetName}" creada exitosamente.`);
    }
    
    const headers = sheet.getRange(1, 1, 1, 5).getValues()[0];
    if (headers[0] !== "ID_Detalle") {
      sheet.getRange("A1:E1").setValues([
        ["ID_Detalle", "ID_Movimiento", "ID_Producto", "ID_Pedido", "Monto_Producto"]
      ]);
      sheet.setFrozenRows(1);
      sheet.getRange("A:E").applyRowBanding();
      ui.alert(`Cabeceras de "${sheetName}" configuradas.`);
    } else {
      ui.alert(`La hoja "${sheetName}" ya está configurada.`);
    }
  } catch (e) {
    ui.alert(`Error en la configuración: ${e.message}`);
  }
}

function abrirForm() {
  const html = HtmlService.createHtmlOutputFromFile('form')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function doGet(e) {
  var email = Session.getActiveUser().getEmail();
  if (email.toLowerCase() === ADMIN_EMAIL.toLowerCase()) {
    const html = HtmlService.createHtmlOutputFromFile('form')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("Registro de Finanzas"); 
    return html;
  } else {
    return HtmlService.createHtmlOutput(
      `<h1>Acceso Denegado</h1><p>No tienes permiso para ver esta aplicación.</p>`
    ).setTitle("Acceso Denegado");
  }
}

// =============================================================================
// 3. OPERACIONES DE DATOS (Lectura)
// =============================================================================

/**
 * ⚡ FUNCIÓN PRINCIPAL DE CARGA (Usada para el startup y la recarga automática)
 */
function getInitialData() {
  try {
    const { categories, metodos, movimientosIds } = _getCategoriasYMetodos(); 
    const { pedidosMap, productosMap } = _getPedidosYProductos(); 
    
    return {
      categories: categories,
      metodos: metodos, 
      movimientosIds: movimientosIds, 
      pedidosMap: pedidosMap,
      productosMap: productosMap 
    };
  } catch (e) {
    Logger.log(e);
    throw new Error(`Error al cargar datos: ${e.message}`);
  }
}

function _getCategoriasYMetodos() {
  const ss = SpreadsheetApp.openById(FINANZAS_SS_ID);
  const sh = ss.getSheetByName("data");
  if (!sh) throw new Error("Hoja 'data' no encontrada.");
  const shMov = ss.getSheetByName("Movimientos"); 
  if (!shMov) throw new Error("Hoja 'Movimientos' no encontrada.");

  // Lógica para Categorías, Métodos y MovimientosIds
  const range = sh.getDataRange();
  const values = range.getValues();
  const headers = values[0]; 

  const tipoCol = headers.findIndex(h => h.toLowerCase().trim() === 'tipo');
  const catCol = headers.findIndex(h => h.toLowerCase().trim() === 'nombre_categoria');
  const metodoCol = headers.findIndex(h => h.toLowerCase().trim() === 'metodo_pago'); 
  
  const metodosDisponibles = metodoCol !== -1;
  const categoryMap = {};
  const metodos = [];
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    const tipo = row[tipoCol];
    const categoria = row[catCol];
    if (tipo && categoria) {
      if (tipo === "Ingreso" || tipo === "Gasto") {
        if (!categoryMap[tipo]) {
          categoryMap[tipo] = [];
        }
        categoryMap[tipo].push(categoria);
      }
    }
    if (metodosDisponibles) {
        const metodo = row[metodoCol];
        if (metodo && !metodos.includes(metodo)) { 
          metodos.push(metodo);
        }
    }
  }
  
  // Lógica para IDs de Movimiento
  const movValues = shMov.getDataRange().getValues();
  const idMovCol = 0; // Columna A: ID_Movimiento
  const movimientosIds = [];

  for (let i = 1; i < movValues.length; i++) {
    const id = movValues[i][idMovCol];
    if (id) movimientosIds.push(String(id).trim());
  }

  return { categories: categoryMap, metodos: metodos, movimientosIds: movimientosIds };
}

/**
 * Busca el costo en el libro externo 'SkusPasados'.
 */
function _searchSkuCostInRespaldo(skuToFind) {
  if (!skuToFind) return null;
  try {
    const ss = SpreadsheetApp.openById(SKUS_PASADOS_SS_ID);
    const sheet = ss.getSheetByName(HOJA_SKUS_PASADOS);
    if (!sheet) {
      Logger.log("ERROR: Hoja " + HOJA_SKUS_PASADOS + " no encontrada en SkusPasados.");
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const normalizedSkuToFind = String(skuToFind).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowSku = String(row[RES_SKU_COL]).trim().toLowerCase();
      
      if (rowSku === normalizedSkuToFind) {
        return Number(row[RES_COSTO_COL]) || 0;
      }
    }
    return null;
  } catch (e) {
    Logger.log("Error al buscar en SkusPasados: " + e.toString());
    return null;
  }
}

/**
 * Obtiene la suma de pagos previos realizados por cada ID Producto.
 */
function _getPagosPrevios() {
  try {
    const ss = SpreadsheetApp.openById(FINANZAS_SS_ID);
    const shDetalle = ss.getSheetByName("Movimiento_Detalle");
    if (!shDetalle || shDetalle.getLastRow() < 2) return new Map();

    const data = shDetalle.getRange(2, 3, shDetalle.getLastRow() - 1, 3).getValues(); 
    const pagosMap = new Map();

    data.forEach(row => {
      const idProducto = String(row[0]).trim();
      const monto = Number(row[2]) || 0; 
      
      if (idProducto) {
        const totalActual = pagosMap.get(idProducto) || 0;
        pagosMap.set(idProducto, totalActual + monto);
      }
    });

    return pagosMap;

  } catch (e) {
    Logger.log("Error en _getPagosPrevios: " + e.message);
    return new Map();
  }
}

/**
 * Incluye la lógica de costo de respaldo y lee todos los productos y pedidos.
 */
function _getPedidosYProductos() {
  try {
    const ss = SpreadsheetApp.openById(PEDIDOS_SS_ID);
    const sh = ss.getSheetByName(HOJA_DE_PEDIDOS);
    if (!sh) throw new Error(`Hoja '${HOJA_DE_PEDIDOS}' no encontrada.`);
    const pagosPrevios = _getPagosPrevios();
    const range = sh.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    
    const cols = {
      pedido: headers.findIndex(h => h.toLowerCase().trim() === COLUMNA_ID_PEDIDO.toLowerCase().trim()),
      producto: headers.findIndex(h => h.toLowerCase().trim() === COLUMNA_ID_PRODUCTO.toLowerCase().trim()),
      cliente: headers.findIndex(h => h.toLowerCase().trim() === COLUMNA_CLIENTE.toLowerCase().trim()),
      sku: headers.findIndex(h => h.toLowerCase().trim() === COLUMNA_SKU.toLowerCase().trim()),
      marcaModelo: headers.findIndex(h => h.toLowerCase().trim() === COLUMNA_MARCA_MODELO.toLowerCase().trim()), 
      precio: headers.findIndex(h => h.toLowerCase().trim() === COLUMNA_PRECIO.toLowerCase().trim()),
      costo: headers.findIndex(h => h.toLowerCase().trim() === COLUMNA_COSTO.toLowerCase().trim())
    };
    
    for (const key in cols) {
      if (cols[key] === -1 && key !== 'costo') throw new Error(`Columna de '${key}' no encontrada.`); 
    }
    
    const pedidosMap = new Map();
    const productosMap = new Map(); 
    const productosVistosEnPedido = new Set(); 
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const idPedido = row[cols.pedido];
      const idProducto = row[cols.producto];
      const productoSku = row[cols.sku];
      const uniqueKey = `${idPedido}_${idProducto}`; 
      
      if (idPedido && idProducto) {
        
        let costo = cols.costo !== -1 ? (Number(row[cols.costo]) || 0) : 0;
        let precio = Number(row[cols.precio]) || 0;
        
        if (costo === 0 && productoSku) {
             const respaldoCosto = _searchSkuCostInRespaldo(productoSku);
             if (respaldoCosto !== null && respaldoCosto > 0) { 
                 costo = respaldoCosto;
             }
        }
        
        const pagado = pagosPrevios.get(idProducto) || 0; 
        const saldoPendiente = costo - pagado; 

        const producto = {
          idProd: idProducto,
          idPed: idPedido, 
          cliente: row[cols.cliente],
          sku: productoSku,
          modeloMarca: row[cols.marcaModelo], 
          precio: precio,
          costo: costo,
          pagado: pagado, 
          pendiente: Math.max(0, saldoPendiente) 
        };
        
        if (!productosVistosEnPedido.has(uniqueKey)) {
          if (!pedidosMap.has(idPedido)) {
            pedidosMap.set(idPedido, []);
          }
          pedidosMap.get(idPedido).push(producto);
          productosVistosEnPedido.add(uniqueKey); 
        }
        
        if (!productosMap.has(idProducto)) {
           productosMap.set(idProducto, producto);
        }
      }
    }
    return {
      pedidosMap: Object.fromEntries(pedidosMap),
      productosMap: Object.fromEntries(productosMap) 
    };
  } catch (e) {
    Logger.log(e);
    throw new Error(`Error al leer Pedidos/Productos: ${e.message}`);
  }
}

/**
 * ⚡ ¡NUEVA FUNCIÓN! Obtiene el detalle de un movimiento para edición.
 */
function getDetalleMovimiento(idMovimiento) {
    if (!idMovimiento) return null;
    try {
        const ss = SpreadsheetApp.openById(FINANZAS_SS_ID);
        const shDetalle = ss.getSheetByName("Movimiento_Detalle");
        const shMov = ss.getSheetByName("Movimientos");
        if (!shDetalle || shDetalle.getLastRow() < 2) return null;
        if (!shMov || shMov.getLastRow() < 2) return null;

        // Leer Movimientos para obtener la fila principal y Metodo/Comprobante
        const movData = shMov.getDataRange().getValues();
        const movHeaders = movData[0];
        const movCols = {
            id: 0, // Columna A
            monto: movHeaders.findIndex(h => h.toLowerCase().trim() === 'monto'),
            metodo: movHeaders.findIndex(h => h.toLowerCase().trim() === 'metodo pago'),
            comprobante: movHeaders.findIndex(h => h.toLowerCase().trim() === 'comprobante'),
            categoria: movHeaders.findIndex(h => h.toLowerCase().trim() === 'categoría')
        };
        const mainMovRow = movData.find(row => String(row[movCols.id]).trim() === String(idMovimiento).trim());
        if (!mainMovRow) return null;

        const categoria = String(mainMovRow[movCols.categoria]).toLowerCase().trim();
        if (categoria !== 'pago pedido') return { error: "Solo se puede editar el detalle de 'Pago Pedido'." };


        // Leer Movimiento_Detalle
        const detalleData = shDetalle.getDataRange().getValues();
        const detalleHeaders = detalleData[0];
        const detalleCols = {
            idDetalle: 0, // Columna A
            idMov: 1, // Columna B
            idProd: 2, // Columna C
            montoProd: 4 // Columna E
        };

        const lineasDetalle = [];
        for (let i = 1; i < detalleData.length; i++) {
            const row = detalleData[i];
            if (String(row[detalleCols.idMov]).trim() === String(idMovimiento).trim()) {
                lineasDetalle.push({
                    rowIndex: i + 1, // Fila real en Detalle
                    idProd: String(row[detalleCols.idProd]).trim(),
                    monto: Number(row[detalleCols.montoProd]) || 0
                });
            }
        }
        
        // Complementar con datos de Producto para mostrar detalles (Cliente, Costo Total)
        // Se llama a _getPedidosYProductos para tener el saldo actual
        const { productosMap } = _getPedidosYProductos();
        
        const detallesCompletos = lineasDetalle.map(detalle => {
            const prodData = productosMap[detalle.idProd] || {};
            // El pagado previo debe excluir el monto de esta misma línea de movimiento que estamos editando
            const pagadoPrevio = (prodData.pagado || 0) - detalle.monto;
            
            return {
                ...detalle,
                cliente: prodData.cliente || 'N/A',
                costoTotal: prodData.costo || 0,
                modeloMarca: prodData.modeloMarca || 'N/A',
                pagadoPrevio: pagadoPrevio,
                pendienteFinal: prodData.costo - pagadoPrevio // Lo que queda pendiente antes de aplicar el monto actual de la línea
            };
        });

        return {
            metodo: String(mainMovRow[movCols.metodo]).trim(),
            comprobante: String(mainMovRow[movCols.comprobante]).trim(),
            lineas: detallesCompletos,
            montoTotalActual: Number(mainMovRow[movCols.monto]) || 0
        };

    } catch (e) {
        Logger.log("Error en getDetalleMovimiento: " + e.message);
        throw new Error("Error al obtener detalle del movimiento: " + e.message);
    }
}

/**
 * ⚡ ¡NUEVA FUNCIÓN! Actualiza las líneas de pago y recalcula el monto principal.
 */
function actualizarPagos(data) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(15000);
    } catch (e) {
        throw new Error("El servidor está ocupado.");
    }

    try {
        const ss = SpreadsheetApp.openById(FINANZAS_SS_ID);
        const shDetalle = ss.getSheetByName("Movimiento_Detalle");
        const shMov = ss.getSheetByName("Movimientos");
        
        let nuevoMontoTotal = 0;
        let idsActualizados = [];

        // 1. Actualizar líneas de Detalle y calcular el nuevo total
        data.lineas.forEach(linea => {
            if (linea.monto < 0) throw new Error("El monto de pago no puede ser negativo.");
            
            const montoActualizado = Number(linea.monto);
            shDetalle.getRange(linea.rowIndex, 5).setValue(montoActualizado); // Columna E (Monto_Producto)
            nuevoMontoTotal += montoActualizado;
            idsActualizados.push(linea.idProd);
        });

        // 2. Actualizar Metodo y Comprobante en Movimientos (Hoja Principal)
        const movData = shMov.getDataRange().getValues();
        const movHeaders = movData[0];
        const movCols = {
            id: 0,
            monto: movHeaders.findIndex(h => h.toLowerCase().trim() === 'monto'),
            metodo: movHeaders.findIndex(h => h.toLowerCase().trim() === 'metodo pago'),
            comprobante: movHeaders.findIndex(h => h.toLowerCase().trim() === 'comprobante'),
        };
        const movRowIndex = movData.findIndex(row => String(row[movCols.id]).trim() === String(data.idMovimiento).trim());

        if (movRowIndex === -1) throw new Error("ID de Movimiento principal no encontrado.");

        const mainRow = shMov.getRange(movRowIndex + 1, 1, 1, shMov.getLastColumn());
        const mainRowValues = mainRow.getValues()[0];
        
        mainRowValues[movCols.monto] = nuevoMontoTotal;
        mainRowValues[movCols.metodo] = data.metodo;
        mainRowValues[movCols.comprobante] = data.comprobante;

        mainRow.setValues([mainRowValues]);

        // 3. Re-verificar estado de los productos (CRÍTICO)
        if (idsActualizados.length > 0) {
            const { productosMap } = _getPedidosYProductos(); 
            let productosAEnviar = [];
            let todosPagados = true;
            
            for (const idProd of idsActualizados) {
                const productoFinal = productosMap[idProd];
                // El productoFinal ya tiene los saldos actualizados después de las re-lecturas y escritura.
                if (productoFinal && productoFinal.pendiente > 0.1) {
                    todosPagados = false;
                }
                if (productoFinal) {
                    productosAEnviar.push(idProd);
                }
            }

            // El estado solo se cambia si todos los productos en la transacción están 100% cubiertos.
            const nuevoEstado = todosPagados ? "En despacho" : "Pendiente";
            _actualizarEstadoEnPedidos(productosAEnviar, nuevoEstado);
        }

        lock.releaseLock();
        
        // 4. Devolver la data inicial para la recarga automática del frontend
        const updatedData = getInitialData();

        return { 
          status: "success", 
          nuevoMonto: nuevoMontoTotal,
          updatedData: updatedData
        };

    } catch (e) {
        lock.releaseLock();
        throw new Error(`Error al actualizar pagos: ${e.message}`);
    }
}

/**
 * ⚡ ¡NUEVA FUNCIÓN! Revisa si un comprobante ya existe.
 */
function checkComprobanteExists(comprobante) {
    if (!comprobante || String(comprobante).trim().length < 3) return false;
    try {
        const ss = SpreadsheetApp.openById(FINANZAS_SS_ID);
        const shMov = ss.getSheetByName("Movimientos");
        if (!shMov || shMov.getLastRow() < 2) return false;

        const movData = shMov.getDataRange().getValues();
        const movHeaders = movData[0];
        const compCol = movHeaders.findIndex(h => h.toLowerCase().trim() === 'comprobante');
        
        if (compCol === -1) return false; 

        const normalizedComprobante = String(comprobante).trim().toLowerCase();

        for (let i = 1; i < movData.length; i++) {
            const rowComprobante = String(movData[i][compCol]).trim().toLowerCase();
            if (rowComprobante === normalizedComprobante) {
                return true; 
            }
        }
        return false;
    } catch (e) {
        Logger.log("Error al verificar comprobante: " + e.message);
        return false; 
    }
}


function registrarMovimiento(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
  } catch (e) { 
    throw new Error("El servidor está ocupado. Intenta de nuevo en 15 segundos.");
  }
  
  let actualizacionExitosa = false;
  
  try {
    const ss = SpreadsheetApp.openById(FINANZAS_SS_ID);
    const shMovimientos = ss.getSheetByName("Movimientos");
    const shDetalle = ss.getSheetByName("Movimiento_Detalle"); 
    
    if (!shMovimientos) throw new Error("Hoja 'Movimientos' no encontrada.");
    if (!shDetalle) throw new Error("Hoja 'Movimiento_Detalle' no encontrada. Ejecuta la configuración desde el menú 'Finanzas'.");

    let newId;
    const lastRow = shMovimientos.getLastRow();
    if (lastRow < 2) {
      newId = "FIN-001"; 
    } else {
      const lastIdCell = shMovimientos.getRange(lastRow, 1).getValue(); 
      const lastNum = parseInt(lastIdCell.split('-')[1] || 0);
      newId = "FIN-" + String(lastNum + 1).padStart(3, '0');
    }
    
    _registrarFilaMovimiento(shMovimientos, d, newId);
    
    const categoria = d.categoria.toLowerCase().trim();
    const esVenta = categoria === 'venta';
    const esPagoPedido = categoria === 'pago pedido';

    let productosAProcesar = [];
    
    if (d.productosJSON) { 
      try {
        productosAProcesar = JSON.parse(d.productosJSON);
      } catch(e) { /* no-op */ }
    }
    
    if (productosAProcesar.length > 0) {
      _registrarFilasDetalle(shDetalle, newId, productosAProcesar, esVenta, d.idPedido);
    } else if (categoria === 'pago envio' && d.idPedido) {
      _registrarDetalleEnvio(shDetalle, newId, d.idPedido, d.monto);
    }
    
    // VERIFICACIÓN DE ESTADO (CRÍTICO)
    if ((esVenta || esPagoPedido) && productosAProcesar.length > 0) {
      
      const { productosMap } = _getPedidosYProductos(); 
      
      let productosAEnviar = [];
      let todosPagados = true;

      for (const prod of productosAProcesar) {
          const idProd = prod.id;
          const productoFinal = productosMap[idProd];
          
          if (!productoFinal) continue;
          
          if (productoFinal.pendiente > 0.1) {
              todosPagados = false;
              break; 
          }
          productosAEnviar.push(idProd);
      }

      if (todosPagados && productosAEnviar.length > 0) {
        _actualizarEstadoEnPedidos(productosAEnviar, "En despacho");
        actualizacionExitosa = true;
      }
    }
    
    lock.releaseLock(); 
    
    // ⚡ PUNTO CRÍTICO: Devolver la data para la recarga automática
    const updatedData = getInitialData();
    
    return { 
      status: 'success', 
      newId: newId, 
      actualizadoPedidos: actualizacionExitosa,
      updatedData: updatedData
    };
    
  } catch (e) {
    lock.releaseLock(); 
    Logger.log(e);
    throw new Error(`Error en el registro: ${e.message}`);
  }
}

function _registrarFilaMovimiento(sheet, d, newId) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = new Array(headers.length);
  headers.forEach((header, index) => {
    const headerNorm = header.toLowerCase().trim().replace(/_/g, " "); 
    if (headerNorm === "id movimiento") newRow[index] = newId;
    else if (headerNorm === "fecha") newRow[index] = d.fecha;
    else if (headerNorm === "tipo") newRow[index] = d.tipo;
    else if (headerNorm === "categoría" || headerNorm === "categoria") newRow[index] = d.categoria;
    else if (headerNorm === "monto") newRow[index] = d.monto;
    else if (headerNorm === "descripción" || headerNorm === "descripcion") newRow[index] = d.descripcion;
    // else if (headerNorm === "sku producto") // Ignorado
    else if (headerNorm === "id pedido") newRow[index] = d.idPedido;
    else if (headerNorm === "modalidad") newRow[index] = d.tipoPago; 
    else if (headerNorm === "metodo pago") newRow[index] = d.metodoPago; 
    else if (headerNorm === "id productos pagados") newRow[index] = d.productosJSON; 
    else if (headerNorm === "comprobante") newRow[index] = d.comprobante; 
  });
  sheet.appendRow(newRow);
}

function _registrarFilasDetalle(sheet, idMovimiento, productos, esVenta, idPedidoPago) {
  const lastDetalleRow = sheet.getLastRow();
  const filasNuevas = [];
  productos.forEach((prod, index) => {
    const idDetalle = `${idMovimiento}-${index + 1}`; 
    const idProducto = prod.id;
    const idPedido = esVenta ? prod.idPedido : idPedidoPago; 
    
    const monto = esVenta ? prod.monto : (prod.costo || prod.monto); 
    
    filasNuevas.push([
      idDetalle,
      idMovimiento,
      idProducto,
      idPedido,
      monto
    ]);
  });
  if (filasNuevas.length > 0) {
    sheet.getRange(lastDetalleRow + 1, 1, filasNuevas.length, 5).setValues(filasNuevas);
  }
}

function _registrarDetalleEnvio(sheet, idMovimiento, idPedido, montoTotal) {
  try {
    const { pedidosMap } = _getPedidosYProductos();
    const productos = pedidosMap[idPedido];
    if (!productos || productos.length === 0) {
      Logger.log(`No se encontraron productos para el Pedido ${idPedido} al prorratear envío.`);
      return; 
    }
    const cantidad = productos.length;
    const montoPorProducto = montoTotal / cantidad;
    const lastDetalleRow = sheet.getLastRow();
    const filasNuevas = [];
    productos.forEach((prod, index) => {
      const idDetalle = `${idMovimiento}-${index + 1}`;
      filasNuevas.push([
        idDetalle,
        idMovimiento,
        prod.idProd,
        idPedido,
        montoPorProducto
      ]);
    });
    if (filasNuevas.length > 0) {
      sheet.getRange(lastDetalleRow + 1, 1, filasNuevas.length, 5).setValues(filasNuevas);
    }
  } catch (e) {
    Logger.log(`Error en _registrarDetalleEnvio: ${e.message}`);
  }
}

function _actualizarEstadoEnPedidos(idProductos, nuevoEstado) {
  if (!idProductos || idProductos.length === 0) return;
  try {
    const ss = SpreadsheetApp.openById(PEDIDOS_SS_ID);
    const sh = ss.getSheetByName(HOJA_DE_PEDIDOS);
    if (!sh) throw new Error("Hoja 'Pedidos' no encontrada.");
    const range = sh.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    const prodCol = headers.findIndex(h => h.toLowerCase().trim() === COLUMNA_ID_PRODUCTO.toLowerCase().trim());
    const estadoCol = headers.findIndex(h => h.toLowerCase().trim() === COLUMNA_ESTADO.toLowerCase().trim());
    if (prodCol === -1) throw new Error(`Columna '${COLUMNA_ID_PRODUCTO}' no encontrada.`);
    if (estadoCol === -1) throw new Error(`Columna '${COLUMNA_ESTADO}' no encontrada.`);
    let cambiosRealizados = 0;
    const idSet = new Set(idProductos);
    for (let i = 1; i < values.length; i++) {
      if (idSet.has(values[i][prodCol])) {
        values[i][estadoCol] = nuevoEstado;
        cambiosRealizados++;
      }
    }
    if (cambiosRealizados > 0) {
      range.setValues(values);
      Logger.log(`Se actualizaron ${cambiosRealizados} filas en 'Pedidos' a '${nuevoEstado}'.`);
    }
  } catch (e) {
    Logger.log(`Error al actualizar estado en Pedidos: ${e.message}`);
  }
}
