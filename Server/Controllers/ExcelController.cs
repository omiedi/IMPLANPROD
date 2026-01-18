using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Server.DTOs;
using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/excel")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;
        private readonly ExcelService _excelService;

        public ExcelController(ILogger<ExcelController> logger, ExcelService excelService)
        {
            _logger = logger;
            _excelService = excelService;
        }

        [HttpPost("exportar-masterlistado")]
        public IActionResult ExportarMasterListado([FromBody] List<MasterListadoIndexDTO> datos)
        {
            try
            {
                var archivo = _excelService.GenerarExcel(datos, "Productos");
                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "Reporte_Productos.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de MasterListado");
                return StatusCode(500, "Error interno al generar Excel.");
            }
        }

        /// <summary>
        /// Exporta remitos (RemitoAnularDTO) a Excel
        /// </summary>
        /// <param name="datos">Lista de remitos a exportar</param>
        /// <returns>Archivo Excel con los remitos</returns>
        [HttpPost("exportar-remitos")]
        public IActionResult ExportarRemitos([FromBody] List<IMPLANPROD.Shared.DTOs.RemitoAnularDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "DoPriCor", "Remito" },
                    { "NumeClie", "Cliente" },
                    { "RazonSocial", "Razón Social" },
                    { "FechMov", "Fecha" },
                    { "ReferCor", "Referencia" }
                };

                var orden = new List<string>
                {
                    "DoPriCor", "NumeClie", "RazonSocial", "FechMov", "ReferCor"
                };

                var archivo = _excelService.GenerarExcel(datos, "Remitos", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Remitos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Remitos");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }
        //movimientos ficha kardex
        [HttpPost("exportar-movimientosKardex")]
        public IActionResult ExportarMovimientos([FromBody] List<MovimientoHistoricoDTO> datos)
        {
            var encabezados = new Dictionary<string, string>
                {

                 { "Fecha", "Fecha" },
                 { "TipoMovimiento", "Tipo" },
                 { "DocumentoNumero", "Documento" },
                 { "Transaccion", "Transaccion" },
                 { "Deposito", "Deposito" },
                 { "IdBulto", "IdBulto" },
                 { "Referencia", "Referencia" },
                 { "Entrada", "Entrada" },
                 { "Salida", "Salida" },
                 { "Saldo", "Saldo" },
                 { "CodInte", "CodInte" }

                };

            var orden = new List<string> { "Fecha", "TipoMovimiento", "DocumentoNumero", "Transaccion", "Deposito", "IdBulto", "Referencia", "Entrada",
                "Salida", "Saldo","CodInte"};

            var archivo = _excelService.GenerarExcel(datos, "Historial", encabezados, orden);

            return File(archivo,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Historial_Movimientos.xlsx");
        }

        // Método de generación con ClosedXML
        private byte[] GenerarExcel(List<MasterListadoIndexDTO> datos)
        {
            using var workbook = new XLWorkbook();
            var hoja = workbook.Worksheets.Add("Productos");

            hoja.Cell(1, 1).Value = "Id";
            hoja.Cell(1, 2).Value = "Código Interno";
            hoja.Cell(1, 3).Value = "Código";
            hoja.Cell(1, 4).Value = "Descripción";
            hoja.Cell(1, 5).Value = "U. Medida";
            hoja.Cell(1, 6).Value = "Tipo Material";

            for (int i = 0; i < datos.Count; i++)
            {
                var fila = i + 2;
                hoja.Cell(fila, 1).Value = datos[i].Id;
                hoja.Cell(fila, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                hoja.Cell(fila, 2).Value = datos[i].Codint;
                hoja.Cell(fila, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                hoja.Cell(fila, 3).Value = datos[i].Codigo;
                hoja.Cell(fila, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                hoja.Cell(fila, 4).Value = datos[i].Descripcion;
                hoja.Cell(fila, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                hoja.Cell(fila, 5).Value = datos[i].Umedida;
                hoja.Cell(fila, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                hoja.Cell(fila, 6).Value = datos[i].TipoMaterialDescripcion;
                hoja.Cell(fila, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            }

            hoja.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }
        [HttpPost("exportar-masterlistado-avanzado")]
        public IActionResult ExportarMasterListadoAvanzado([FromBody] List<MasterListadoIndexDTO> datos)
        {
            var encabezados = new Dictionary<string, string>
    {
        { "Id", "ID" },
        { "Codint", "Código Interno" },
        { "Codigo", "Código" },
        { "Descripcion", "Descripción" },
        { "Umedida", "U. Medida" },
        { "TipoMaterialDescripcion", "Tipo Material" }
    };

            var orden = new List<string>
    {
        "Id", "Codigo", "Codint", "Descripcion", "Umedida", "TipoMaterialDescripcion"
            };

            var archivo = _excelService.GenerarExcel(datos, "Productos", encabezados, orden);

            return File(archivo,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Reporte_Productos.xlsx");
        }

        /// <summary>
        /// Exporta pedidos de venta a Excel
        /// </summary>
        /// <param name="datos">Lista de pedidos a exportar</param>
        /// <returns>Archivo Excel con los pedidos</returns>
        [HttpPost("exportar-pedidos")]
        public IActionResult ExportarPedidos([FromBody] List<Pedenca> datos)
        {
            try
            {
                var encabezados = new Dictionary<string, string>
                {
                    { "NroPed", "N° Pedido" },
                    { "FechEmis", "Fecha Emisión" },
                    { "RasoClie", "Cliente" },
                    { "NroClie", "N° Cliente" },
                    { "NroOcom", "N° OC" },
                    { "Status", "Estado" },
                    { "FeSolEnt", "Fecha Entrega" },
                    { "VenRep", "Vendedor" },
                    { "Vended", "Cod. Vendedor" },
                    { "Referencia", "Referencia" },
                    { "TranspEnt", "Transporte" },
                    { "NTranspo", "N° Transporte" },
                    { "ListPre", "Lista Precio" },
                    { "CondPag", "Cond. Pago" },
                    { "PorComVend", "% Comisión" }
                };

                var orden = new List<string>
                {
                    "NroPed", "FechEmis", "RasoClie", "NroClie", "NroOcom", 
                    "Status", "FeSolEnt", "VenRep", "Vended", "Referencia",
                    "TranspEnt", "NTranspo", "ListPre", "CondPag", "PorComVend"
                };

                var archivo = _excelService.GenerarExcel(datos, "Pedidos de Venta", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Pedidos_Venta_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Pedidos de Venta");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        /// <summary>
        /// Exporta órdenes de compra a Excel
        /// </summary>
        /// <param name="datos">Lista de órdenes de compra a exportar</param>
        /// <returns>Archivo Excel con las órdenes de compra</returns>
        [HttpPost("exportar-ordenes-compra")]
        public IActionResult ExportarOrdenesCompra([FromBody] List<Ocomenca> datos)
        {
            try
            {
                var encabezados = new Dictionary<string, string>
                {
                    { "NumeOcom", "N° Orden" },
                    { "FemiOcom", "Fecha Emisión" },
                    { "NproOcom", "N° Proveedor" },
                    { "RaSoProv", "Razón Social" },
                    { "RefeOcom", "Referencia" },
                    { "StatOcom", "Estado" },
                    { "TotalOcom", "Total" },
                    { "FGenOcom", "Fecha Generación" },
                    { "MoneEmis", "Moneda" },
                    { "CondPag", "Cond. Pago" },
                    { "Observa", "Observaciones" }
                };

                var orden = new List<string>
                {
                    "NumeOcom", "FemiOcom", "NproOcom", "RaSoProv", "RefeOcom", 
                    "StatOcom", "TotalOcom", "FGenOcom", "MoneEmis", "CondPag", "Observa"
                };

                var archivo = _excelService.GenerarExcel(datos, "Órdenes de Compra", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Ordenes_Compra_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Órdenes de Compra");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        /// <summary>
        /// Exporta detalle de orden de compra a Excel
        /// </summary>
        /// <param name="datos">Lista de items del detalle de la orden</param>
        /// <returns>Archivo Excel con el detalle de la orden</returns>
        [HttpPost("exportar-detalle-orden")]
        public IActionResult ExportarDetalleOrden([FromBody] List<Ocomdeta> datos)
        {
            try
            {
                var encabezados = new Dictionary<string, string>
                {
                    { "NuOrdItem", "Item" },
                    { "CodExte", "Código" },
                    { "Descritem", "Descripción" },
                    { "CantOcom", "Cantidad" },
                    { "CantRec", "Recibido" },
                    { "PreUnit", "Precio Unit." },
                    { "INetItem", "Total" }
                };

                var orden = new List<string>
                {
                    "NuOrdItem", "CodExte", "Descritem", "CantOcom", 
                    "CantRec", "PreUnit", "INetItem"
                };

                var archivo = _excelService.GenerarExcel(datos, "Detalle de Orden de Compra", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Detalle_Orden_Compra_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Detalle de Orden de Compra");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        /// <summary>
        /// Exporta órdenes pendientes del proveedor a Excel
        /// </summary>
        /// <param name="datos">Lista de órdenes pendientes del proveedor</param>
        /// <returns>Archivo Excel con las órdenes pendientes</returns>
        [HttpPost("exportar-ordenes-pendientes")]
        public IActionResult ExportarOrdenesPendientes([FromBody] List<Ocomdeta> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarOrdenesPendientes] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarOrdenesPendientes] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "NumeOcom", "N° Orden" },
                    { "NuOrdItem", "Item" },
                    { "CodExte", "Código" },
                    { "Descritem", "Descripción" },
                    { "CantOcom", "Cantidad" },
                    { "CantRec", "Recibido" },
                    { "FemiOcom", "Fecha Emisión" }
                };

                var orden = new List<string>
                {
                    "NumeOcom", "NuOrdItem", "CodExte", "Descritem", 
                    "CantOcom", "CantRec", "FemiOcom"
                };

                _logger.LogInformation($"[ExportarOrdenesPendientes] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Órdenes Pendientes del Proveedor", encabezados, orden);

                _logger.LogInformation($"[ExportarOrdenesPendientes] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Ordenes_Pendientes_Proveedor_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarOrdenesPendientes] Error al generar Excel de Órdenes Pendientes del Proveedor");
                _logger.LogError($"[ExportarOrdenesPendientes] Detalles: {ex.Message}");
                _logger.LogError($"[ExportarOrdenesPendientes] StackTrace: {ex.StackTrace}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta materiales en consignación a Excel
        /// </summary>
        /// <param name="datos">Lista de materiales en consignación</param>
        /// <returns>Archivo Excel con los materiales en consignación</returns>
        [HttpPost("exportar-materiales-consignacion")]
        public IActionResult ExportarMaterialesConsignacion([FromBody] List<Maconsig> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarMaterialesConsignacion] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarMaterialesConsignacion] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "FemiOcom", "Fecha Emisión" },
                    { "NproOcom", "N° Proveedor" },
                    { "RaSoProv", "Razón Social" },
                    { "NroOrfa", "N° Orden Fab." },
                    { "NroOcom", "N° Orden Compra" },
                    { "CodInte", "Código Interno" },
                    { "CodExte", "Código Externo" },
                    { "Descritem", "Descripción" },
                    { "Unidad", "Unidad" },
                    { "CantAsig", "Cant. Asignada" },
                    { "CantEntreg", "Cant. Entregada" },
                    { "FechaEntreg", "Fecha Entrega" },
                    { "CantRendid", "Cant. Rendida" },
                    { "SaldoPend", "Saldo Pendiente" }
                };

                var orden = new List<string>
                {
                    "FemiOcom", "NproOcom", "RaSoProv", "NroOrfa", "NroOcom",
                    "CodInte", "CodExte", "Descritem", "Unidad",
                    "CantAsig", "CantEntreg", "FechaEntreg", "CantRendid", "SaldoPend"
                };

                _logger.LogInformation($"[ExportarMaterialesConsignacion] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Materiales en Consignación", encabezados, orden);

                _logger.LogInformation($"[ExportarMaterialesConsignacion] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Materiales_Consignacion_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarMaterialesConsignacion] Error al generar Excel de Materiales en Consignación");
                _logger.LogError($"[ExportarMaterialesConsignacion] Detalles: {ex.Message}");
                _logger.LogError($"[ExportarMaterialesConsignacion] StackTrace: {ex.StackTrace}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta recepciones por rango de fechas a Excel
        /// </summary>
        /// <param name="datos">Lista de recepciones</param>
        /// <returns>Archivo Excel con las recepciones</returns>
        [HttpPost("exportar-recepciones-rango")]
        public IActionResult ExportarRecepcionesRango([FromBody] List<Bolrecep> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarRecepcionesRango] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarRecepcionesRango] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "NumeBoRe", "N° Boleta" },
                    { "Descrip", "Descripción" },
                    { "IdenLote", "Lote" },
                    { "FechRecep", "Fecha Recep." },
                    { "CodInte", "Cód. Interno" },
                    { "Unidad", "Unidad" },
                    { "CantNom", "Cantidad" },
                    { "DepoEnt", "Depósito" },
                    { "NumProv", "N° Prov." },
                    { "LoteColProv", "Lote Prov." },
                    { "NumRemProv", "Remito Prov." },
                    { "CodExte", "Cód. Externo" },
                    { "NumOC", "N° O.C." },
                    { "MoneVaUni", "Moneda" },
                    { "ValoUnit", "Valor Unit." },
                    { "CanToApro", "Cant. Aprob." },
                    { "CanToRech", "Cant. Rech." },
                    { "CanToFalta", "Cant. Falta" },
                    { "FechLibe", "Fecha Liber." },
                    { "Observa", "Observaciones" }
                };

                var orden = new List<string>
                {
                    "NumeBoRe", "Descrip", "IdenLote", "FechRecep", "CodInte",
                    "Unidad", "CantNom", "DepoEnt", "NumProv", "LoteColProv",
                    "NumRemProv", "CodExte", "NumOC", "MoneVaUni", "ValoUnit",
                    "CanToApro", "CanToRech", "CanToFalta", "FechLibe", "Observa"
                };

                _logger.LogInformation($"[ExportarRecepcionesRango] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Recepciones por Rango de Fechas", encabezados, orden);

                _logger.LogInformation($"[ExportarRecepcionesRango] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Recepciones_Rango_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarRecepcionesRango] Error al generar Excel de Recepciones");
                _logger.LogError($"[ExportarRecepcionesRango] Detalles: {ex.Message}");
                _logger.LogError($"[ExportarRecepcionesRango] StackTrace: {ex.StackTrace}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta recepciones con rechazos a Excel
        /// </summary>
        /// <param name="datos">Lista de recepciones con rechazos</param>
        /// <returns>Archivo Excel con las recepciones con rechazos</returns>
        [HttpPost("exportar-recepciones-rechazos")]
        public IActionResult ExportarRecepcionesRechazos([FromBody] List<Bolrecep> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarRecepcionesRechazos] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarRecepcionesRechazos] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "NumeBoRe", "N° Boleta" },
                    { "Descrip", "Descripción" },
                    { "IdenLote", "Lote" },
                    { "FechRecep", "Fecha Recep." },
                    { "CodInte", "Cód. Interno" },
                    { "Unidad", "Unidad" },
                    { "CantNom", "Cantidad" },
                    { "DepoEnt", "Depósito" },
                    { "NumProv", "N° Prov." },
                    { "LoteColProv", "Lote Prov." },
                    { "NumRemProv", "Remito Prov." },
                    { "NumFacProv", "Factura Prov." },
                    { "NumOC", "N° O.C." },
                    { "MoneVaUni", "Moneda" },
                    { "ValoUnit", "Valor Unit." },
                    { "CanToApro", "Cant. Aprob." },
                    { "CanToRech", "Cant. Rech." },
                    { "CanToFalta", "Cant. Falta" },
                    { "FechLibe", "Fecha Liber." },
                    { "Observa", "Observaciones" }
                };

                var orden = new List<string>
                {
                    "NumeBoRe", "Descrip", "IdenLote", "FechRecep", "CodInte",
                    "Unidad", "CantNom", "DepoEnt", "NumProv", "LoteColProv",
                    "NumRemProv", "NumFacProv", "NumOC", "MoneVaUni", "ValoUnit",
                    "CanToApro", "CanToRech", "CanToFalta", "FechLibe", "Observa"
                };

                _logger.LogInformation($"[ExportarRecepcionesRechazos] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Recepciones con Rechazos", encabezados, orden);

                _logger.LogInformation($"[ExportarRecepcionesRechazos] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Recepciones_Rechazos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarRecepcionesRechazos] Error al generar Excel de Recepciones con Rechazos");
                _logger.LogError($"[ExportarRecepcionesRechazos] Detalles: {ex.Message}");
                _logger.LogError($"[ExportarRecepcionesRechazos] StackTrace: {ex.StackTrace}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta recepciones sin orden de compra a Excel
        /// </summary>
        /// <param name="datos">Lista de recepciones sin orden de compra</param>
        /// <returns>Archivo Excel con las recepciones sin OC</returns>
        [HttpPost("exportar-recepciones-sin-oc")]
        public IActionResult ExportarRecepcionesSinOC([FromBody] List<Bolrecep> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarRecepcionesSinOC] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarRecepcionesSinOC] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "NumeBoRe", "N° Boleta" },
                    { "Descrip", "Descripción" },
                    { "IdenLote", "Lote" },
                    { "FechRecep", "Fecha Recep." },
                    { "CodInte", "Cód. Interno" },
                    { "Unidad", "Unidad" },
                    { "CantNom", "Cantidad" },
                    { "DepoEnt", "Depósito" },
                    { "NumProv", "N° Prov." },
                    { "LoteColProv", "Lote Prov." },
                    { "NumRemProv", "Remito Prov." },
                    { "NumFacProv", "Factura Prov." },
                    { "NumOC", "N° O.C." },
                    { "MoneVaUni", "Moneda" },
                    { "ValoUnit", "Valor Unit." },
                    { "CanToApro", "Cant. Aprob." },
                    { "CanToRech", "Cant. Rech." },
                    { "CanToFalta", "Cant. Falta" },
                    { "FechLibe", "Fecha Liber." },
                    { "Observa", "Observaciones" }
                };

                var orden = new List<string>
                {
                    "NumeBoRe", "Descrip", "IdenLote", "FechRecep", "CodInte",
                    "Unidad", "CantNom", "DepoEnt", "NumProv", "LoteColProv",
                    "NumRemProv", "NumFacProv", "NumOC", "MoneVaUni", "ValoUnit",
                    "CanToApro", "CanToRech", "CanToFalta", "FechLibe", "Observa"
                };

                _logger.LogInformation($"[ExportarRecepcionesSinOC] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Recepciones sin Orden de Compra", encabezados, orden);

                _logger.LogInformation($"[ExportarRecepcionesSinOC] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Recepciones_Sin_OC_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarRecepcionesSinOC] Error al generar Excel de Recepciones sin OC");
                _logger.LogError($"[ExportarRecepcionesSinOC] Detalles: {ex.Message}");
                _logger.LogError($"[ExportarRecepcionesSinOC] StackTrace: {ex.StackTrace}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta listado de recepción a Excel
        /// </summary>
        /// <param name="datos">Lista de diccionarios con los datos a exportar</param>
        /// <returns>Archivo Excel con el listado de recepción</returns>
        [HttpPost("exportar-recepcion")]
        public IActionResult ExportarRecepcion([FromBody] List<Dictionary<string, object?>> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarRecepcion] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarRecepcion] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                // Log del primer registro para debugging
                var primerRegistro = datos.First();
                _logger.LogInformation($"[ExportarRecepcion] Primer registro tiene {primerRegistro.Keys.Count} columnas: {string.Join(", ", primerRegistro.Keys)}");

                // Los datos ya vienen con los encabezados traducidos como claves
                // No necesitamos diccionario de encabezados personalizados
                var orden = datos.First().Keys.ToList();

                _logger.LogInformation($"[ExportarRecepcion] Generando Excel con {datos.Count} filas y {orden.Count} columnas");

                var archivo = _excelService.GenerarExcel(datos, "Recepción de Órdenes", null, orden);

                _logger.LogInformation($"[ExportarRecepcion] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Recepcion_Ordenes_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarRecepcion] Error al generar Excel de Recepción");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        /// <summary>
        /// Exporta capacidad de máquinas/equipos a Excel
        /// </summary>
        /// <param name="datos">Lista de capacidad de máquinas</param>
        /// <returns>Archivo Excel con la capacidad de máquinas</returns>
        [HttpPost("exportar-capacidad-maquinas")]
        public IActionResult ExportarCapacidadMaquinas([FromBody] List<IMPLANPROD.Shared.DTOs.CapacidadMaquinaDTO> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarCapacidadMaquinas] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarCapacidadMaquinas] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "CodigoMaquina", "Código Máquina" },
                    { "DescripcionMaquina", "Descripción Máquina/Equipo" },
                    { "TiempoFormateado", "Tiempo Requerido (hh:mm:ss)" }
                };

                var orden = new List<string>
                {
                    "CodigoMaquina", "DescripcionMaquina", "TiempoFormateado"
                };

                _logger.LogInformation($"[ExportarCapacidadMaquinas] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Capacidad Máquinas", encabezados, orden);

                _logger.LogInformation($"[ExportarCapacidadMaquinas] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Capacidad_Maquinas_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarCapacidadMaquinas] Error al generar Excel de Capacidad de Máquinas");
                _logger.LogError($"[ExportarCapacidadMaquinas] Detalles: {ex.Message}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta capacidad de mano de obra directa (M.O.D.) a Excel
        /// </summary>
        /// <param name="datos">Lista de capacidad de mano de obra</param>
        /// <returns>Archivo Excel con la capacidad de M.O.D.</returns>
        [HttpPost("exportar-capacidad-mano-obra")]
        public IActionResult ExportarCapacidadManoObra([FromBody] List<IMPLANPROD.Shared.DTOs.CapacidadManoObraDTO> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarCapacidadManoObra] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarCapacidadManoObra] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "Categoria", "Categoría" },
                    { "DescripcionCategoria", "Descripción" },
                    { "TiempoFormateado", "Tiempo Requerido (hh:mm:ss)" }
                };

                var orden = new List<string>
                {
                    "Categoria", "DescripcionCategoria", "TiempoFormateado"
                };

                _logger.LogInformation($"[ExportarCapacidadManoObra] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Capacidad M.O.D.", encabezados, orden);

                _logger.LogInformation($"[ExportarCapacidadManoObra] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Capacidad_ManoObra_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarCapacidadManoObra] Error al generar Excel de Capacidad de M.O.D.");
                _logger.LogError($"[ExportarCapacidadManoObra] Detalles: {ex.Message}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta detalle de programa de fabricación a Excel
        /// </summary>
        /// <param name="datos">Lista de detalles del programa</param>
        /// <returns>Archivo Excel con el detalle del programa</returns>
        [HttpPost("exportar-detalle-programa")]
        public IActionResult ExportarDetallePrograma([FromBody] List<IMPLANPROD.Shared.DTOs.DeprofabDTO> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarDetallePrograma] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarDetallePrograma] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "CodigoExterno", "Código" },
                    { "Descripcion", "Descripción Producto" },
                    { "Cantidad", "Cantidad" },
                    { "Observaciones", "Observaciones" }
                };

                var orden = new List<string>
                {
                    "CodigoExterno", "Descripcion", "Cantidad", "Observaciones"
                };

                _logger.LogInformation($"[ExportarDetallePrograma] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Detalle Programa", encabezados, orden);

                _logger.LogInformation($"[ExportarDetallePrograma] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Detalle_Programa_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarDetallePrograma] Error al generar Excel de Detalle de Programa");
                _logger.LogError($"[ExportarDetallePrograma] Detalles: {ex.Message}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta detalle de piezas por programa a Excel
        /// </summary>
        /// <param name="datos">Lista de piezas con cantidades por programa</param>
        /// <returns>Archivo Excel con el detalle de piezas</returns>
        [HttpPost("exportar-detalle-piezas")]
        public IActionResult ExportarDetallePiezas([FromBody] List<IMPLANPROD.Shared.DTOs.PiezaProgramaDto> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarDetallePiezas] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarDetallePiezas] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "CodigoProducto", "Código" },
                    { "DescripcionPieza", "Pieza" },
                    { "Promedio", "Promedio" }
                };

                var orden = new List<string>
                {
                    "CodigoProducto", "DescripcionPieza", "Promedio"
                };

                _logger.LogInformation($"[ExportarDetallePiezas] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Detalle Piezas", encabezados, orden);

                _logger.LogInformation($"[ExportarDetallePiezas] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Detalle_Piezas_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarDetallePiezas] Error al generar Excel de Detalle de Piezas");
                _logger.LogError($"[ExportarDetallePiezas] Detalles: {ex.Message}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta carga horaria de máquinas y operarios a Excel
        /// </summary>
        /// <param name="datos">Lista de carga horaria por máquina y operario</param>
        /// <returns>Archivo Excel con la carga horaria</returns>
        [HttpPost("exportar-carga-horaria-maquinas")]
        public IActionResult ExportarCargaHorariaMaquinas([FromBody] List<IMPLANPROD.Shared.DTOs.CargaHorariaMaquinaDto> datos)
        {
            try
            {
                _logger.LogInformation($"[ExportarCargaHorariaMaquinas] Recibiendo solicitud con {datos?.Count ?? 0} registros");

                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarCargaHorariaMaquinas] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "DescripcionMaquina", "Máquina" },
                    { "CodigoOperacion", "Cód-Oper" },
                    { "NombreOperario", "Operador" },
                    { "Promedio", "Promedio (hs)" }
                };

                var orden = new List<string>
                {
                    "DescripcionMaquina", "CodigoOperacion", "NombreOperario", "Promedio"
                };

                _logger.LogInformation($"[ExportarCargaHorariaMaquinas] Generando Excel con {datos.Count} filas");

                var archivo = _excelService.GenerarExcel(datos, "Carga Horaria", encabezados, orden);

                _logger.LogInformation($"[ExportarCargaHorariaMaquinas] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Carga_Horaria_Maquinas_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarCargaHorariaMaquinas] Error al generar Excel de Carga Horaria");
                _logger.LogError($"[ExportarCargaHorariaMaquinas] Detalles: {ex.Message}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

    }
}