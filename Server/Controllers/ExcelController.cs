using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Shared.DTOs;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
using IMPLANPROD.Server.Data;
using System.IO;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/excel")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;
        private readonly ExcelService _excelService;
        private readonly IWebHostEnvironment _env;
        private readonly DataContext _db;

        public ExcelController(ILogger<ExcelController> logger, ExcelService excelService, IWebHostEnvironment env, DataContext db)
        {
            _logger = logger;
            _excelService = excelService;
            _env = env;
            _db = db;
        }

        [HttpPost("exportar-orden-compra")]
        public IActionResult ExportarOrdenCompra([FromBody] IMPLANPROD.Shared.DTOs.OrdenCompraConDetalleDTO datos)
        {
            try
            {
                if (datos?.Encabezado == null)
                {
                    return BadRequest("No se recibió el encabezado de la orden de compra");
                }

                if (datos.Detalle == null || !datos.Detalle.Any())
                {
                    return BadRequest("No hay items para exportar");
                }

                var encabezado = datos.Encabezado;
                var items = datos.Detalle
                    .Where(i => (i.StatOcom ?? 0) < 10)
                    .OrderBy(i => i.NuOrdItem)
                    .ToList();

                var reemplazarCodigoPropio = false;
                var flex = _db.Set<FLEXCRL>().FirstOrDefault(f => f.CLABUS == "OCOMPRA" && f.PROBUS == "RECOPROP");
                if (flex != null && int.TryParse(flex.VALCONTROL, out var valFlex) && valFlex > 0)
                {
                    reemplazarCodigoPropio = true;
                }

                var revisarCodigo = false;
                var flexRevi = _db.Set<FLEXCRL>().FirstOrDefault(f => f.CLABUS == "OCOMPRA" && f.PROBUS == "REVICOD");
                if (flexRevi != null && int.TryParse(flexRevi.VALCONTROL, out var valFlexRevi) && valFlexRevi > 0)
                {
                    revisarCodigo = true;
                }

                Dictionary<int, string> equivalenciasPorCodInte = new();
                if (reemplazarCodigoPropio && encabezado.NproOcom.HasValue && encabezado.NproOcom.Value > 0)
                {
                    equivalenciasPorCodInte = _db.Replacos
                        .AsNoTracking()
                        .Where(r => r.NumeProv == encabezado.NproOcom.Value && r.CodInte > 0 && r.Referencia != null && r.Referencia != "")
                        .GroupBy(r => r.CodInte)
                        .ToDictionary(g => g.Key, g => g.First().Referencia!.Trim());
                }

                Dictionary<int, (string? Revision, DateTime? Ferevis)> revisionesPorCodInte = new();
                if (revisarCodigo)
                {
                    var codints = items
                        .Select(i => i.CodInte)
                        .Where(ci => ci.HasValue && ci.Value > 0)
                        .Select(ci => (int)ci!.Value)
                        .Distinct()
                        .ToList();

                    revisionesPorCodInte = _db.masters
                        .AsNoTracking()
                        .Where(m => codints.Contains(m.Codint))
                        .Select(m => new { m.Codint, m.Revision, m.Ferevis })
                        .ToDictionary(x => x.Codint, x => (x.Revision, x.Ferevis));
                }

                var nro = encabezado.NumeOcom ?? 0;
                var version = (short)((encabezado.OcomRevis ?? 0) + 1);
                var nroConVersion = $"{nro:000000}-{version:000}";

                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("OrdenCompra");

                // Config general
                ws.Style.Font.FontName = "Arial";
                ws.Style.Font.FontSize = 10;
                ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
                ws.PageSetup.PageOrientation = XLPageOrientation.Portrait;

                // Anchos aproximados para parecerse al impreso
                ws.Column(1).Width = 6;   // Item
                ws.Column(2).Width = 18;  // Código
                ws.Column(3).Width = 42;  // Descripción
                ws.Column(4).Width = 6;   // UM
                ws.Column(5).Width = 12;  // Cantidad
                ws.Column(6).Width = 12;  // Precio
                ws.Column(7).Width = 14;  // Importe
                ws.Column(8).Width = 12;  // F. Entrega

                var empresaActiva = _db.Empresas
                    .Where(e => e.status == 1)
                    .OrderBy(e => e.Id)
                    .FirstOrDefault();

                var fila = 1;

                // Encabezado empresa (logo + datos)
                ws.Row(fila).Height = 45;
                ws.Row(fila + 1).Height = 15;
                ws.Row(fila + 2).Height = 15;
                ws.Row(fila + 3).Height = 15;
                ws.Row(fila + 4).Height = 15;

                var logoPath = Path.Combine(_env.ContentRootPath, "..", "Client", "wwwroot", "images", "logoempresa.png");
                if (System.IO.File.Exists(logoPath))
                {
                    using var logoStream = new MemoryStream(System.IO.File.ReadAllBytes(logoPath));
                    ws.AddPicture(logoStream)
                        .MoveTo(ws.Cell(fila, 1))
                        .WithSize(140, 40);
                }

                if (empresaActiva != null)
                {
                    var nombreProvincia = "";
                    if (empresaActiva.provinciaemp.HasValue)
                    {
                        nombreProvincia = _db.States
                            .Where(s => s.Id == empresaActiva.provinciaemp.Value)
                            .Select(s => s.Name)
                            .FirstOrDefault() ?? "";
                    }

                    var direccion = $"{empresaActiva.calle} {empresaActiva.numero} {empresaActiva.complemento}".Trim();
                    var localidad = (empresaActiva.localidad ?? "").Trim();
                    var web = (empresaActiva.WEB ?? "").Trim();
                    var tel = (empresaActiva.TEL1 ?? "").Trim();

                    ws.Range(fila + 1, 1, fila + 1, 4).Merge();
                    ws.Range(fila + 2, 1, fila + 2, 4).Merge();
                    ws.Range(fila + 3, 1, fila + 3, 4).Merge();

                    ws.Cell(fila + 1, 1).Value = $"Dirección: {direccion} - {localidad},";
                    ws.Cell(fila + 2, 1).Value = $"({nombreProvincia} - Tel.: {tel})";
                    ws.Cell(fila + 3, 1).Value = web;

                    ws.Range(fila + 1, 1, fila + 3, 4).Style.Font.FontSize = 9;
                    ws.Range(fila + 1, 1, fila + 3, 4).Style.Alignment.WrapText = true;
                }

                fila += 6;

                // Encabezado (título)
                ws.Cell(fila, 1).Value = "PEDIDO DE SUMINISTRO";
                ws.Range(fila, 1, fila, 8).Merge();
                ws.Range(fila, 1, fila, 8).Style.Font.Bold = true;
                ws.Range(fila, 1, fila, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(fila, 1, fila, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                fila++;

                ws.Cell(fila, 1).Value = "Nro:";
                ws.Cell(fila, 2).Value = nroConVersion;
                ws.Cell(fila, 6).Value = "Fecha:";
                ws.Cell(fila, 7).Value = encabezado.FemiOcom?.ToString("dd/MM/yyyy") ?? "";
                ws.Range(fila, 1, fila, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(fila, 1, fila, 8).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                ws.Range(fila, 2, fila, 2).Style.Font.Bold = true;
                fila += 2;

                // Proveedor
                ws.Cell(fila, 1).Value = "Proveedor:";
                ws.Cell(fila, 2).Value = $"{encabezado.NproOcom} - {encabezado.RaSoProv}";
                ws.Range(fila, 2, fila, 8).Merge();
                fila++;

                ws.Cell(fila, 1).Value = "Condición Pago:";
                ws.Cell(fila, 2).Value = encabezado.CpagOcom?.ToString(CultureInfo.InvariantCulture) ?? "";
                ws.Cell(fila, 4).Value = "Moneda:";
                ws.Cell(fila, 5).Value = encabezado.MoneEmis?.ToString(CultureInfo.InvariantCulture) ?? "";
                ws.Range(fila - 1, 1, fila, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(fila - 1, 1, fila, 8).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                fila += 2;

                // Tabla encabezados
                var filaHeader = fila;
                ws.Cell(fila, 1).Value = "Item";
                ws.Cell(fila, 2).Value = "Código";
                ws.Cell(fila, 3).Value = "Descripción";
                ws.Cell(fila, 4).Value = "UM";
                ws.Cell(fila, 5).Value = "Cantidad";
                ws.Cell(fila, 6).Value = "Precio";
                ws.Cell(fila, 7).Value = "Importe";
                ws.Cell(fila, 8).Value = "F. Entrega";
                ws.Range(fila, 1, fila, 8).Style.Font.Bold = true;
                ws.Range(fila, 1, fila, 8).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Range(fila, 1, fila, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(fila, 1, fila, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(fila, 1, fila, 8).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                fila++;

                var filaInicioItems = fila;

                foreach (var it in items)
                {
                    ws.Cell(fila, 1).Value = it.NuOrdItem ?? 0;
                    var codigoImpreso = it.CodExte ?? "";
                    if (reemplazarCodigoPropio && it.CodInte.HasValue)
                    {
                        var key = (int)it.CodInte.Value;
                        if (equivalenciasPorCodInte.TryGetValue(key, out var codEq) && !string.IsNullOrWhiteSpace(codEq))
                        {
                            codigoImpreso = codEq;
                        }
                    }

                    if (revisarCodigo && it.CodInte.HasValue)
                    {
                        var key = (int)it.CodInte.Value;
                        if (revisionesPorCodInte.TryGetValue(key, out var revData))
                        {
                            var rv = (revData.Revision ?? string.Empty).Trim();
                            if (!string.IsNullOrWhiteSpace(rv) && revData.Ferevis.HasValue)
                            {
                                var ferevx = revData.Ferevis.Value.ToString("dd/MM/yy");
                                codigoImpreso = $"{codigoImpreso} (Letra: {rv} - {ferevx})";
                            }
                        }
                    }

                    ws.Cell(fila, 2).Value = codigoImpreso;
                    ws.Cell(fila, 3).Value = it.Descritem ?? "";
                    ws.Cell(fila, 4).Value = it.UniMas ?? "";
                    ws.Cell(fila, 5).Value = it.CantOcom ?? 0m;
                    ws.Cell(fila, 6).Value = it.PreUnit ?? 0m;
                    ws.Cell(fila, 8).Value = it.FentreSol?.ToString("dd/MM/yy") ?? "";

                    // Formatos numéricos
                    ws.Cell(fila, 5).Style.NumberFormat.Format = "0.0";
                    ws.Cell(fila, 6).Style.NumberFormat.Format = "0.00";
                    ws.Cell(fila, 7).Style.NumberFormat.Format = "0.00";

                    // Importe = Cantidad * Precio (editable)
                    ws.Cell(fila, 7).FormulaA1 = $"=E{fila}*F{fila}";

                    // Bordes
                    ws.Range(fila, 1, fila, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(fila, 1, fila, 8).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    // Alineaciones
                    ws.Cell(fila, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(fila, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(fila, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(fila, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(fila, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    fila++;
                }

                var filaFinItems = fila - 1;

                // Totales
                fila++;
                ws.Cell(fila, 6).Value = "TOTAL";
                ws.Cell(fila, 6).Style.Font.Bold = true;
                ws.Cell(fila, 7).FormulaA1 = $"=SUM(G{filaInicioItems}:G{filaFinItems})";
                ws.Cell(fila, 7).Style.Font.Bold = true;
                ws.Cell(fila, 7).Style.NumberFormat.Format = "0.00";
                ws.Range(fila, 6, fila, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(fila, 6, fila, 7).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Congelar encabezados de tabla
                ws.SheetView.FreezeRows(filaHeader);

                using var stream = new MemoryStream();
                wb.SaveAs(stream);
                var bytes = stream.ToArray();

                return File(bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"OrdenCompra_{nroConVersion}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Orden de Compra");
                return StatusCode(500, "Error interno al generar Excel.");
            }
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
        /// Exporta la lista de componentes (Formulas) de un producto a Excel.
        /// Usado desde /master/navegacion-componentes con el botón Imprimir/Exp.Excel.
        /// </summary>
        [HttpPost("exportar-componentes")]
        public IActionResult ExportarComponentes([FromBody] List<Formulas> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "CodiElem", "Códint" },
                    { "CodXElem", "Código" },
                    { "MasterDescripcion", "Descripción" },
                    { "UsoBruto", "Cantidad" },
                    { "UMED_FORM", "Unidad" }
                };

                var orden = new List<string>
                {
                    "CodiElem", "CodXElem", "MasterDescripcion", "UsoBruto", "UMED_FORM"
                };

                var archivo = _excelService.GenerarExcel(datos, "Componentes", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Componentes_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Componentes");
                return StatusCode(500, "Error al generar el archivo excel");
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

        [HttpPost("exportar-estructura-componentes")]
        public IActionResult ExportarEstructuraComponentes([FromBody] List<DetalleStockDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                    return BadRequest("No hay datos para exportar");

                var encabezados = new Dictionary<string, string>
                {
                    { "ItemJerarquico", "ItemJerarquico" },
                    { "Nivel",       "Nivel"       },
                    { "CI1",         "CodiElem"    },
                    { "CodXEelem",   "CodXEelem"   },
                    { "Descripcion", "Descripción" },
                    { "UM",          "U.M."        },
                    { "Uso",         "Uso"         }
                };

                var orden = new List<string> { "ItemJerarquico", "Nivel", "CI1", "CodXEelem", "Descripcion", "UM", "Uso" };

                var archivo = _excelService.GenerarExcel(datos, "Estructura Componentes", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Estructura_Componentes_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Estructura Componentes");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        [HttpPost("exportar-stock-componentes")]
        public IActionResult ExportarStockComponentes([FromBody] List<DetalleStockDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var depositos = datos
                    .SelectMany(d => d.StockPorDeposito?.Keys ?? Enumerable.Empty<string>())
                    .Where(k => !string.IsNullOrWhiteSpace(k))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(k => k)
                    .ToList();

                var encabezados = new Dictionary<string, string>
                {
                    { "CI1", "CI1" },
                    { "CodXEelem", "CodXEelem" },
                    { "Descripcion", "Descripción" },
                    { "UM", "U.M." },
                    { "Uso", "Uso" },
                    { "CantSimulada", "Cant.Simulada" },
                    { "StAcumulado", "St.Acumulado" },
                };

                var orden = new List<string>
                {
                    "CI1",
                    "CodXEelem",
                    "Descripcion",
                    "UM",
                    "Uso",
                    "CantSimulada",
                    "StAcumulado",
                };

                foreach (var dep in depositos)
                {
                    var key = $"Dep.{dep}";
                    encabezados[key] = $"Dep. {dep}";
                    orden.Add(key);
                }

                var datosPlano = new List<Dictionary<string, object?>>();
                foreach (var item in datos)
                {
                    var row = new Dictionary<string, object?>
                    {
                        ["CI1"] = item.CI1,
                        ["CodXEelem"] = item.CodXEelem,
                        ["Descripcion"] = item.Descripcion,
                        ["UM"] = item.UM,
                        ["Uso"] = item.Uso,
                        ["CantSimulada"] = item.CantSimulada,
                        ["StAcumulado"] = item.StAcumulado,
                    };

                    foreach (var dep in depositos)
                    {
                        var key = $"Dep.{dep}";
                        if (item.StockPorDeposito != null && item.StockPorDeposito.TryGetValue(dep, out var stock))
                        {
                            row[key] = stock;
                        }
                        else
                        {
                            row[key] = 0m;
                        }
                    }

                    datosPlano.Add(row);
                }

                var archivo = _excelService.GenerarExcel(datosPlano, "Stock Componentes", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Stock_Componentes_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Stock Componentes");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        [HttpPost("exportar-depositos-posibles-saldo")]
        public IActionResult ExportarDepositosPosiblesSaldo([FromBody] List<DepositosPosiblesConSaldoDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "Codigo", "Código" },
                    { "Descripcion", "Descripción" },
                    { "SaldoStock", "Saldo Stock" },
                    { "Deposito", "Depósito" }
                };

                var orden = new List<string>
                {
                    "Codigo",
                    "Descripcion",
                    "SaldoStock",
                    "Deposito"
                };

                var archivo = _excelService.GenerarExcel(datos, "Depósitos Posibles con Saldo", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Depositos_Posibles_Saldo_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Depósitos Posibles con Saldo");
                return StatusCode(500, "Error al generar el archivo excel");
            }
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

        [HttpPost("exportar-pedidos-pendientes-detalle")]
        public IActionResult ExportarPedidosPendientesDetalle([FromBody] List<PedidoVentaPendienteDetalleDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { nameof(PedidoVentaPendienteDetalleDTO.NroClie), "N° Cliente" },
                    { nameof(PedidoVentaPendienteDetalleDTO.Rasoclie), "Razón Social" },
                    { nameof(PedidoVentaPendienteDetalleDTO.Codigo), "Código" },
                    { nameof(PedidoVentaPendienteDetalleDTO.Descripcion), "Descripción" },
                    { nameof(PedidoVentaPendienteDetalleDTO.NroPed), "N° Pedido" },
                    { nameof(PedidoVentaPendienteDetalleDTO.CantSolicitada), "Cant. Solicitada" },
                    { nameof(PedidoVentaPendienteDetalleDTO.UMedida), "U. Medida" },
                    { nameof(PedidoVentaPendienteDetalleDTO.CantEntregada), "Cant. Entregada" },
                    { nameof(PedidoVentaPendienteDetalleDTO.Saldo), "Saldo" },
                    { nameof(PedidoVentaPendienteDetalleDTO.FechaEntrega), "Fecha Entrega" }
                };

                var orden = new List<string>
                {
                    nameof(PedidoVentaPendienteDetalleDTO.NroClie),
                    nameof(PedidoVentaPendienteDetalleDTO.Rasoclie),
                    nameof(PedidoVentaPendienteDetalleDTO.Codigo),
                    nameof(PedidoVentaPendienteDetalleDTO.Descripcion),
                    nameof(PedidoVentaPendienteDetalleDTO.NroPed),
                    nameof(PedidoVentaPendienteDetalleDTO.CantSolicitada),
                    nameof(PedidoVentaPendienteDetalleDTO.UMedida),
                    nameof(PedidoVentaPendienteDetalleDTO.CantEntregada),
                    nameof(PedidoVentaPendienteDetalleDTO.Saldo),
                    nameof(PedidoVentaPendienteDetalleDTO.FechaEntrega)
                };

                var archivo = _excelService.GenerarExcel(datos, "Pedidos Pendientes", encabezados, orden);

                return File(archivo,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"Pedidos_Pendientes_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Pedidos Pendientes (detalle)");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        [HttpPost("exportar-pedidos-entregas-detalle")]
        public IActionResult ExportarPedidosEntregasDetalle([FromBody] List<PedidoVentaEntregaDetalleDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { nameof(PedidoVentaEntregaDetalleDTO.NroPed), "N° Pedido" },
                    { nameof(PedidoVentaEntregaDetalleDTO.Fecha), "Fecha" },
                    { nameof(PedidoVentaEntregaDetalleDTO.Cliente), "Cliente" },
                    { nameof(PedidoVentaEntregaDetalleDTO.Codigo), "Código" },
                    { nameof(PedidoVentaEntregaDetalleDTO.Descripcion), "Descripción" },
                    { nameof(PedidoVentaEntregaDetalleDTO.CanPed), "Cant. Pedida" },
                    { nameof(PedidoVentaEntregaDetalleDTO.CantEnt), "Cant. Entregada" },
                    { nameof(PedidoVentaEntregaDetalleDTO.Saldo), "Saldo" },
                    { nameof(PedidoVentaEntregaDetalleDTO.Remito), "N° Remito" },
                    { nameof(PedidoVentaEntregaDetalleDTO.Estado), "Estado" }
                };

                var orden = new List<string>
                {
                    nameof(PedidoVentaEntregaDetalleDTO.NroPed),
                    nameof(PedidoVentaEntregaDetalleDTO.Fecha),
                    nameof(PedidoVentaEntregaDetalleDTO.Cliente),
                    nameof(PedidoVentaEntregaDetalleDTO.Codigo),
                    nameof(PedidoVentaEntregaDetalleDTO.Descripcion),
                    nameof(PedidoVentaEntregaDetalleDTO.CanPed),
                    nameof(PedidoVentaEntregaDetalleDTO.CantEnt),
                    nameof(PedidoVentaEntregaDetalleDTO.Remito),
                    nameof(PedidoVentaEntregaDetalleDTO.Saldo),
                    nameof(PedidoVentaEntregaDetalleDTO.Estado)
                };

                var archivo = _excelService.GenerarExcel(datos, "Entregas de Pedidos", encabezados, orden);

                return File(archivo,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"Entregas_Pedidos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Entregas de Pedidos (detalle)");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        [HttpPost("exportar-pedidos-con-ocom-por-fechas")]
        public IActionResult ExportarPedidosConOcomPorFechas([FromBody] List<PedidoVentaConOcomDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { nameof(PedidoVentaConOcomDTO.FechEmis), "Fecha Emisión" },
                    { nameof(PedidoVentaConOcomDTO.NroPed), "N° Pedido" },
                    { nameof(PedidoVentaConOcomDTO.NroClie), "N° Cliente" },
                    { nameof(PedidoVentaConOcomDTO.RasoClie), "Cliente" },
                    { nameof(PedidoVentaConOcomDTO.NroOcom), "N° O.Compra" }
                };

                var orden = new List<string>
                {
                    nameof(PedidoVentaConOcomDTO.FechEmis),
                    nameof(PedidoVentaConOcomDTO.NroPed),
                    nameof(PedidoVentaConOcomDTO.NroClie),
                    nameof(PedidoVentaConOcomDTO.RasoClie),
                    nameof(PedidoVentaConOcomDTO.NroOcom)
                };

                var archivo = _excelService.GenerarExcel(datos, "Pedidos con O.Compra", encabezados, orden);

                return File(archivo,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"Pedidos_Con_OCompra_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Pedidos con O.Compra por fechas");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        /// <summary>
        /// Exporta implosión de productos a Excel
        /// </summary>
        /// <param name="datos">Lista de implosión a exportar</param>
        /// <returns>Archivo Excel con la implosión</returns>
        [HttpPost("exportar-implosion")]
        public IActionResult ExportarImplosion([FromBody] List<ImplosionItemDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarImplosion] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "Codigo", "Código" },
                    { "Descripcion", "Descripción" },
                    { "Umedida", "U.Med." },
                    { "UsoBruto", "Uso Bruto" }
                };

                var orden = new List<string>
                {
                    "Codigo",
                    "Descripcion",
                    "Umedida",
                    "UsoBruto"
                };

                var archivo = _excelService.GenerarExcel(datos, "Implosión", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Implosion_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Implosión");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        /// <summary>
        /// Exporta estadística de pedidos de venta a Excel
        /// </summary>
        /// <param name="datos">Lista de detalle de pedidos a exportar</param>
        /// <returns>Archivo Excel con la estadística de pedidos</returns>
        [HttpPost("exportar-estadistica-pedidos")]
        public IActionResult ExportarEstadisticaPedidos([FromBody] List<DetallePedidoEstadisticaDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarEstadisticaPedidos] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "NroPed", "N° Pedido" },
                    { "FechEmis", "Fecha Emisión" },
                    { "NroClie", "N° Cliente" },
                    { "RasoClie", "Cliente" },
                    { "CodiExt", "Código" },
                    { "Descrip", "Descripción" },
                    { "CanPed", "Cant. Pedida" },
                    { "CantidadEntregada", "Cant. Entregada" },
                    { "CantidadPendiente", "Cant. Pendiente" },
                    { "FeSolEnt", "Entrega Solicitada" },
                    { "Estado", "Estado" }
                };

                var orden = new List<string>
                {
                    "NroPed",
                    "FechEmis",
                    "NroClie",
                    "RasoClie",
                    "CodiExt",
                    "Descrip",
                    "CanPed",
                    "CantidadEntregada",
                    "CantidadPendiente",
                    "FeSolEnt",
                    "Estado"
                };

                var archivo = _excelService.GenerarExcel(datos, "Estadística de Pedidos", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Estadistica_Pedidos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Estadística de Pedidos");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        /// <summary>
        /// Exporta modelos (Badamode) a Excel
        /// </summary>
        /// <param name="datos">Lista de modelos a exportar</param>
        /// <returns>Archivo Excel con los modelos</returns>
        [HttpPost("exportar-modelos")]
        public IActionResult ExportarModelos([FromBody] List<Badamode> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    _logger.LogWarning("[ExportarModelos] No hay datos para exportar");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "Cod_Mod", "Código" },
                    { "Descripcion", "Descripción" },
                    { "Ubic", "Ubicación" },
                    { "Desc_Mat", "Material" },
                    { "CantPiezas", "Cant. Piezas" },
                    { "FechPro_Con", "Próx. Control" },
                    { "Material", "Cód. Material" },
                    { "Frecuen", "Frecuencia" },
                    { "FechUlt_Con", "Último Control" },
                    { "Cod_Inte", "Cód. Interno" },
                    { "ObservaMod", "Observaciones" }
                };

                var orden = new List<string>
                {
                    "Cod_Mod", "Descripcion", "Ubic", "Desc_Mat", "CantPiezas", 
                    "FechPro_Con", "Material", "Frecuen", "FechUlt_Con", 
                    "Cod_Inte", "ObservaMod"
                };

                _logger.LogInformation($"[ExportarModelos] Generando Excel con {datos.Count} modelos");

                var archivo = _excelService.GenerarExcel(datos, "Modelos", encabezados, orden);

                _logger.LogInformation($"[ExportarModelos] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Modelos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarModelos] Error al generar Excel de Modelos");
                _logger.LogError($"[ExportarModelos] Detalles: {ex.Message}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta proveedores (Proved12) a Excel
        /// </summary>
        /// <param name="datos">Lista de proveedores a exportar</param>
        /// <returns>Archivo Excel con los proveedores</returns>
        [HttpPost("exportar-proveedores")]
        public IActionResult ExportarProveedores([FromBody] List<Proved12> datos)
        {
            _logger.LogInformation("[ExportarProveedores] ===== INICIO EXPORTACIÓN PROVEEDORES =====");
            _logger.LogInformation($"[ExportarProveedores] Datos recibidos: {datos?.Count ?? 0} registros");
            
            try
            {
                if (datos == null)
                {
                    _logger.LogWarning("[ExportarProveedores] datos es NULL");
                    return BadRequest("No se recibieron datos para exportar");
                }
                
                if (!datos.Any())
                {
                    _logger.LogWarning("[ExportarProveedores] La lista de datos está vacía");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "Id", "Nro Proveedor" },
                    { "Raso_Cli", "Razón Social" },
                    { "Domi_Cli", "Domicilio" },
                    { "Tele_Cli", "Teléfono" },
                    { "Mail_Cli", "Mail" },
                    { "Cuit_Cli", "CUIT" },
                    { "Loca_Cli", "Localidad" },
                    { "Cpos_Cli", "Código Postal" }
                };

                var orden = new List<string>
                {
                    "Id", "Raso_Cli", "Domi_Cli", "Tele_Cli", "Mail_Cli", 
                    "Cuit_Cli", "Loca_Cli", "Cpos_Cli"
                };

                _logger.LogInformation($"[ExportarProveedores] Generando Excel con {datos.Count} proveedores");

                var archivo = _excelService.GenerarExcel(datos, "Proveedores", encabezados, orden);

                _logger.LogInformation($"[ExportarProveedores] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Proveedores_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarProveedores] Error al generar Excel de Proveedores");
                _logger.LogError($"[ExportarProveedores] Detalles: {ex.Message}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta órdenes de fabricación a Excel
        /// </summary>
        /// <param name="datos">Lista de órdenes de fabricación a exportar</param>
        /// <returns>Archivo Excel con las órdenes de fabricación</returns>
        [HttpPost("exportar-ordenes-fabricacion")]
        public IActionResult ExportarOrdenesFabricacion([FromBody] List<OrdefabrIndexDTO> datos)
        {
            _logger.LogInformation("[ExportarOrdenesFabricacion] ===== INICIO EXPORTACIÓN ÓRDENES DE FABRICACIÓN =====");
            _logger.LogInformation($"[ExportarOrdenesFabricacion] Datos recibidos: {datos?.Count ?? 0} registros");
            
            try
            {
                if (datos == null)
                {
                    _logger.LogWarning("[ExportarOrdenesFabricacion] datos es NULL");
                    return BadRequest("No se recibieron datos para exportar");
                }
                
                if (!datos.Any())
                {
                    _logger.LogWarning("[ExportarOrdenesFabricacion] La lista de datos está vacía");
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "FechGen", "Fecha" },
                    { "IdSubOr", "Ord.Fab" },
                    { "Cod_Exte", "Código" },
                    { "Descrip", "Descripción" },
                    { "CanProy", "Cant.Proy." },
                    { "CanApro", "Cant.Apro." },
                    { "Id_Orfa", "ID" },
                    { "StatuOrf", "Estado" }
                };

                var orden = new List<string>
                {
                    "FechGen", "IdSubOr", "Cod_Exte", "Descrip", "CanProy", 
                    "CanApro", "Id_Orfa", "StatuOrf"
                };

                _logger.LogInformation($"[ExportarOrdenesFabricacion] Generando Excel con {datos.Count} órdenes");

                var archivo = _excelService.GenerarExcel(datos, "Órdenes de Fabricación", encabezados, orden);

                _logger.LogInformation($"[ExportarOrdenesFabricacion] Archivo generado exitosamente: {archivo.Length} bytes");

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Ordenes_Fabricacion_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarOrdenesFabricacion] Error al generar Excel de Órdenes de Fabricación");
                _logger.LogError($"[ExportarOrdenesFabricacion] Detalles: {ex.Message}");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
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

        /// <summary>
        /// Exporta a Excel el Plan de Inspección y Ensayo mostrado en el modal de
        /// /ingenieria/documentos. Sigue el mismo patrón de integración que el resto del
        /// sistema (por ejemplo /pedidos-venta): el componente &lt;ImpresionLista&gt; del
        /// cliente envía la lista de filas (List&lt;Plainsp&gt;) tal cual las tiene en pantalla
        /// mediante <c>UrlExportarExcel</c>. Acá se arma manualmente con ClosedXML (en vez de
        /// _excelService.GenerarExcel) porque además de la tabla se necesita un encabezado fijo
        /// con Código, Descripción y el título del formulario; el Código interno (Cod_Inte) para
        /// buscar esos datos en Master se toma del primer ítem recibido.
        /// </summary>
        /// <param name="datos">Filas PLAINSP actualmente cargadas en la grilla del modal.</param>
        /// <returns>Archivo Excel con el Plan de Inspección.</returns>
        [HttpPost("exportar-plan-inspeccion")]
        public IActionResult ExportarPlanInspeccion([FromBody] List<Plainsp> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay filas para exportar");
                }

                // El Código/Descripción del encabezado se buscan en Master a partir del
                // Cod_Inte que ya viaja en cada fila PLAINSP (misma clave que Master.Codint).
                var codInte = datos.FirstOrDefault(d => d.Cod_Inte.HasValue)?.Cod_Inte;
                var master = codInte.HasValue
                    ? _db.masters.AsNoTracking().FirstOrDefault(m => m.Codint == codInte.Value)
                    : null;

                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Plan de Inspección");

                ws.Style.Font.FontName = "Arial";
                ws.Style.Font.FontSize = 10;

                ws.Column(1).Width = 30; // Dimensión s/Plano
                ws.Column(2).Width = 24; // Instrumento
                ws.Column(3).Width = 45; // Resultado

                var fila = 1;

                // Título fijo del formulario
                ws.Cell(fila, 1).Value = "Plan de Inspección y Ensayo";
                ws.Range(fila, 1, fila, 3).Merge();
                ws.Range(fila, 1, fila, 3).Style.Font.Bold = true;
                ws.Range(fila, 1, fila, 3).Style.Font.FontSize = 14;
                ws.Range(fila, 1, fila, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                fila += 2;

                // Encabezado: Código y Descripción del producto
                ws.Cell(fila, 1).Value = "Código:";
                ws.Cell(fila, 1).Style.Font.Bold = true;
                ws.Cell(fila, 2).Value = master?.Codigo ?? string.Empty;
                fila++;

                ws.Cell(fila, 1).Value = "Descripción:";
                ws.Cell(fila, 1).Style.Font.Bold = true;
                ws.Cell(fila, 2).Value = master?.Descripcion ?? string.Empty;
                ws.Range(fila, 2, fila, 3).Merge();
                fila += 2;

                // Encabezado de la tabla de detalle
                ws.Cell(fila, 1).Value = "Dimensión s/Plano";
                ws.Cell(fila, 2).Value = "Instrumento";
                ws.Cell(fila, 3).Value = "Resultado";
                ws.Range(fila, 1, fila, 3).Style.Font.Bold = true;
                ws.Range(fila, 1, fila, 3).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Range(fila, 1, fila, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                var filaEncabezadoDetalle = fila;
                fila++;

                // Detalle PLAINSP en el mismo orden que la grilla del modal (NroOrden).
                // Se exporta tal cual, incluyendo filas vacías que el usuario usa como
                // separador visual (mismo criterio que la impresión A4).
                foreach (var item in datos.OrderBy(d => d.NroOrden).ThenBy(d => d.Id))
                {
                    ws.Cell(fila, 1).Value = item.DimSegunPlano ?? string.Empty;
                    ws.Cell(fila, 2).Value = item.Instrumento ?? string.Empty;
                    ws.Cell(fila, 3).Value = item.Observaciones ?? string.Empty;
                    fila++;
                }

                ws.Range(filaEncabezadoDetalle, 1, fila - 1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(filaEncabezadoDetalle, 1, fila - 1, 3).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                using var stream = new MemoryStream();
                wb.SaveAs(stream);

                var nombreArchivo = $"PlanInspeccion_{Safe(master?.Codigo)}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                return File(stream.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            nombreArchivo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarPlanInspeccion] Error al generar Excel del Plan de Inspección");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        private static string Safe(string? value) => string.IsNullOrWhiteSpace(value)
            ? "SinCodigo"
            : string.Concat(value.Split(Path.GetInvalidFileNameChars()));

        /// <summary>
        /// Exporta a Excel el detalle por lote de las recepciones pendientes de aprobación
        /// (modal "Detalle por Lote" de /recepcion/selproveedor). Mismo patrón que el resto
        /// del sistema: el componente ImpresionLista del cliente envía la lista ya filtrada
        /// mediante UrlExportarExcel; acá se usa _excelService.GenerarExcel porque es una
        /// tabla plana, sin encabezado fijo adicional.
        /// </summary>
        /// <param name="datos">Filas de detalle por lote (ya filtradas en el cliente).</param>
        /// <returns>Archivo Excel con el detalle por lote.</returns>
        [HttpPost("exportar-recepciones-pendientes-detalle")]
        public IActionResult ExportarRecepcionesPendientesDetalle([FromBody] List<RecepcionPendienteAprobacionDetalleDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "NumeroRecepcion", "Nro. Recepción" },
                    { "FechaRecepcion", "Fecha" },
                    { "NumeroOrdenCompra", "Nro. OC" },
                    { "NumeroRemito", "Remito" },
                    { "NumeroProveedor", "N. Proveedor" },
                    { "RazonSocial", "R. Social" },
                    { "IdenLote", "Lote" },
                    { "Codigo", "Código" },
                    { "Descripcion", "Descripción" },
                    { "CantidadNominal", "Cantidad" }
                };

                var orden = new List<string>
                {
                    "NumeroRecepcion", "FechaRecepcion", "NumeroOrdenCompra", "NumeroRemito",
                    "NumeroProveedor", "RazonSocial", "IdenLote", "Codigo", "Descripcion", "CantidadNominal"
                };

                var archivo = _excelService.GenerarExcel(datos, "Detalle por Lote", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Recepciones_Pendientes_Detalle_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarRecepcionesPendientesDetalle] Error al generar Excel del detalle por lote");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta anotaciones de pedido (Anotaped) a Excel.
        /// Recibe la lista de registros ANOTAPED desde el cliente (ImpresionLista).
        /// </summary>
        /// <param name="datos">Lista de registros ANOTAPED a exportar</param>
        /// <returns>Archivo Excel con las anotaciones del pedido</returns>
        [HttpPost("exportar-anotaped")]
        public IActionResult ExportarAnotaped([FromBody] List<Anotaped> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "IdAnotacion", "N° Anot." },
                    { "FechaCambio", "Fecha" },
                    { "CodiExt", "Código" },
                    { "Descrip", "Descripción" },
                    { "CanPed", "Cantidad" },
                    { "Usuario", "Usuario" },
                    { "ObservaItem", "Obs. Item" },
                    { "ObservaGral", "Obs. Gral." }
                };

                var orden = new List<string>
                {
                    "IdAnotacion", "FechaCambio", "CodiExt", "Descrip",
                    "CanPed", "Usuario", "ObservaItem", "ObservaGral"
                };

                var archivo = _excelService.GenerarExcel(datos, "Anotaciones del Pedido", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"Anotaciones_Pedido_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ExportarAnotaped] Error al generar Excel de anotaciones de pedido");
                return StatusCode(500, $"Error al generar el archivo excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Exporta producciones propias (PropiaProduccionDTO) a Excel.
        /// Se usa en la pantalla /propia-produccion.
        /// </summary>
        [HttpPost("exportar-propia-produccion")]
        public IActionResult ExportarPropiaProduccion([FromBody] List<PropiaProduccionDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "FechMov", "Fecha" },
                    { "DoPriLar", "Documento" },
                    { "Cod_Exte", "Código" },
                    { "Descrip", "Descripción" },
                    { "CantIngre", "Cantidad" }
                };

                var orden = new List<string>
                {
                    "FechMov", "DoPriLar", "Cod_Exte", "Descrip", "CantIngre"
                };

                var archivo = _excelService.GenerarExcel(datos, "Producción Propia", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"ProduccionPropia_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de Producción Propia");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

        /// <summary>
        /// Exporta consumos de una producción propia (ConsumoProduccionDTO) a Excel.
        /// Se usa en el modal de consumos de la pantalla /propia-produccion.
        /// </summary>
        [HttpPost("exportar-consumos-produccion")]
        public IActionResult ExportarConsumosProduccion([FromBody] List<ConsumoProduccionDTO> datos)
        {
            try
            {
                if (datos == null || !datos.Any())
                {
                    return BadRequest("No hay datos para exportar");
                }

                var encabezados = new Dictionary<string, string>
                {
                    { "Cod_Exte", "Código" },
                    { "Descrip", "Descripción" },
                    { "CantSalid", "Cantidad Consumida" }
                };

                var orden = new List<string>
                {
                    "Cod_Exte", "Descrip", "CantSalid"
                };

                var archivo = _excelService.GenerarExcel(datos, "Consumos de Producción", encabezados, orden);

                return File(archivo,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"ConsumosProduccion_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar Excel de consumos de producción");
                return StatusCode(500, "Error al generar el archivo excel");
            }
        }

    }
}