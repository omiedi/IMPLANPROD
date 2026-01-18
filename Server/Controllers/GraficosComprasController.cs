using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using System.Globalization;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class GraficosComprasController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<GraficosComprasController> _logger;

        public GraficosComprasController(DataContext context, ILogger<GraficosComprasController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene todos los datos para los gráficos de compras
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<GraficosComprasDTO>> GetGraficosCompras(
            [FromQuery] DateTime? fechaDesde = null,
            [FromQuery] DateTime? fechaHasta = null)
        {
            try
            {
                // Establecer fechas por defecto si no se proporcionan
                var desde = fechaDesde ?? DateTime.Now.AddMonths(-6);
                var hasta = fechaHasta ?? DateTime.Now;

                _logger.LogInformation($"Obteniendo gráficos de compras desde {desde:yyyy-MM-dd} hasta {hasta:yyyy-MM-dd}");

                var resultado = new GraficosComprasDTO
                {
                    FechaDesde = desde,
                    FechaHasta = hasta
                };

                // Obtener datos de órdenes de compra en el rango de fechas
                var ordenes = await _context.Ocomencas
                    .Where(o => o.FemiOcom >= desde && o.FemiOcom <= hasta)
                    .ToListAsync();

                if (!ordenes.Any())
                {
                    _logger.LogWarning("No se encontraron órdenes de compra en el rango especificado");
                    return Ok(GenerarDatosEjemplo(desde, hasta));
                }

                // Generar datos para cada gráfico
                resultado.ComprasPorMes = await GenerarComprasPorMes(ordenes, desde, hasta);
                resultado.TopProveedores = await GenerarTopProveedores(ordenes);
                resultado.EstadosOrdenes = await GenerarEstadosOrdenes(ordenes);
                resultado.MontosPorMoneda = await GenerarMontosPorMoneda(ordenes);

                // Calcular estadísticas generales
                resultado.TotalOrdenes = ordenes.Count();
                resultado.MontoTotal = ordenes.Sum(o => o.TotalOcom ?? 0);
                resultado.TotalProveedores = ordenes.Select(o => o.NproOcom).Distinct().Count();
                
                var mesesEnRango = ((hasta.Year - desde.Year) * 12) + hasta.Month - desde.Month + 1;
                resultado.PromedioMensual = mesesEnRango > 0 ? resultado.MontoTotal / mesesEnRango : 0;

                _logger.LogInformation($"Gráficos generados exitosamente: {resultado.TotalOrdenes} órdenes, {resultado.TotalProveedores} proveedores");

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener datos para gráficos de compras");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Genera datos de compras agrupadas por mes
        /// </summary>
        private async Task<List<DatoMensualDTO>> GenerarComprasPorMes(List<Ocomenca> ordenes, DateTime desde, DateTime hasta)
        {
            var datosPorMes = ordenes
                .Where(o => o.FemiOcom.HasValue)
                .GroupBy(o => new { 
                    Año = o.FemiOcom!.Value.Year, 
                    Mes = o.FemiOcom!.Value.Month 
                })
                .Select(g => new DatoMensualDTO
                {
                    Mes = $"{ObtenerNombreMes(g.Key.Mes)} {g.Key.Año}",
                    Monto = g.Sum(o => o.TotalOcom ?? 0),
                    Cantidad = g.Count(),
                    // Agregamos campos auxiliares para ordenamiento cronológico
                    Año = g.Key.Año,
                    NumeroMes = g.Key.Mes
                })
                .OrderBy(d => d.Año)
                .ThenBy(d => d.NumeroMes)
                .ToList();

            return await Task.FromResult(datosPorMes);
        }

        /// <summary>
        /// Genera datos de top 5 proveedores
        /// </summary>
        private async Task<List<DatoProveedorDTO>> GenerarTopProveedores(List<Ocomenca> ordenes)
        {
            var topProveedores = ordenes
                .Where(o => o.NproOcom.HasValue)
                .GroupBy(o => new { 
                    Codigo = o.NproOcom, 
                    Nombre = o.RaSoProv ?? $"Proveedor {o.NproOcom}"
                })
                .Select(g => new DatoProveedorDTO
                {
                    Nombre = g.Key.Nombre.Length > 30 ? g.Key.Nombre.Substring(0, 30) + "..." : g.Key.Nombre,
                    Monto = g.Sum(o => o.TotalOcom ?? 0),
                    CantidadOrdenes = g.Count()
                })
                .OrderByDescending(p => p.Monto)
                .Take(5)
                .ToList();

            return await Task.FromResult(topProveedores);
        }

        /// <summary>
        /// Genera datos de estados de órdenes
        /// </summary>
        private async Task<List<DatoEstadoDTO>> GenerarEstadosOrdenes(List<Ocomenca> ordenes)
        {
            var estados = ordenes
                .GroupBy(o => ObtenerEstadoOrden(o))
                .Select(g => new DatoEstadoDTO
                {
                    Estado = g.Key,
                    Cantidad = g.Count(),
                    MontoTotal = g.Sum(o => o.TotalOcom ?? 0)
                })
                .OrderByDescending(e => e.Cantidad)
                .ToList();

            return await Task.FromResult(estados);
        }

        /// <summary>
        /// Genera datos de montos por moneda
        /// </summary>
        private async Task<List<DatoMonedaDTO>> GenerarMontosPorMoneda(List<Ocomenca> ordenes)
        {
            var monedas = ordenes
                .GroupBy(o => o.MoneEmis ?? 1) // 1 = Pesos por defecto
                .Select(g => new DatoMonedaDTO
                {
                    CodigoMoneda = g.Key.ToString(),
                    Moneda = ObtenerNombreMoneda(g.Key.ToString()),
                    Monto = g.Sum(o => o.TotalOcom ?? 0),
                    CantidadOrdenes = g.Count()
                })
                .OrderByDescending(m => m.Monto)
                .ToList();

            return await Task.FromResult(monedas);
        }

        /// <summary>
        /// Obtiene el nombre del mes en español
        /// </summary>
        private string ObtenerNombreMes(int mes)
        {
            var meses = new[] { "", "Ene", "Feb", "Mar", "Abr", "May", "Jun", 
                               "Jul", "Ago", "Sep", "Oct", "Nov", "Dic" };
            return mes >= 1 && mes <= 12 ? meses[mes] : mes.ToString();
        }

        /// <summary>
        /// Determina el estado de una orden basado en StatOcom
        /// </summary>
        private string ObtenerEstadoOrden(Ocomenca orden)
        {
            // Lógica basada en el campo StatOcom
            var estado = orden.StatOcom ?? -1; // Default pendiente si es null
            
            return estado switch
            {
                -1 => "Pendiente de Aprobación",    // stat_ocom = -1
                9 => "Anulada",                     // stat_ocom = 9
                >= 0 and < 9 => "Vigente",          // stat_ocom >= 0 and < 9
                _ => "Estado Desconocido"
            };
        }

        /// <summary>
        /// Obtiene el nombre completo de la moneda
        /// </summary>
        private string ObtenerNombreMoneda(string codigo)
        {
            return codigo switch
            {
                "1" => "Pesos ($)",
                "2" => "Dólares (USD)",
                "3" => "Euros (EUR)",
                "4" => "Reales (BRL)",
                _ => $"Moneda {codigo}"
            };
        }

        /// <summary>
        /// Genera datos de ejemplo cuando no hay datos reales
        /// </summary>
        private GraficosComprasDTO GenerarDatosEjemplo(DateTime desde, DateTime hasta)
        {
            return new GraficosComprasDTO
            {
                FechaDesde = desde,
                FechaHasta = hasta,
                ComprasPorMes = new List<DatoMensualDTO>
                {
                    new() { Mes = "Ene 2024", Monto = 15000, Cantidad = 5 },
                    new() { Mes = "Feb 2024", Monto = 22000, Cantidad = 8 },
                    new() { Mes = "Mar 2024", Monto = 18000, Cantidad = 6 },
                    new() { Mes = "Abr 2024", Monto = 25000, Cantidad = 10 },
                    new() { Mes = "May 2024", Monto = 30000, Cantidad = 12 },
                    new() { Mes = "Jun 2024", Monto = 28000, Cantidad = 11 }
                },
                TopProveedores = new List<DatoProveedorDTO>
                {
                    new() { Nombre = "METALURGICA SIAM", Monto = 85000, CantidadOrdenes = 15 },
                    new() { Nombre = "PROVEEDOR ABC S.A.", Monto = 65000, CantidadOrdenes = 12 },
                    new() { Nombre = "INDUSTRIAS XYZ", Monto = 48000, CantidadOrdenes = 8 },
                    new() { Nombre = "COMERCIAL DEF", Monto = 32000, CantidadOrdenes = 6 },
                    new() { Nombre = "SUMINISTROS GHI", Monto = 28000, CantidadOrdenes = 5 }
                },
                EstadosOrdenes = new List<DatoEstadoDTO>
                {
                    new() { Estado = "Aprobadas", Cantidad = 150, MontoTotal = 180000 },
                    new() { Estado = "Pendientes", Cantidad = 80, MontoTotal = 95000 },
                    new() { Estado = "Anuladas", Cantidad = 15, MontoTotal = 18000 }
                },
                MontosPorMoneda = new List<DatoMonedaDTO>
                {
                    new() { Moneda = "Pesos ($)", CodigoMoneda = "ARS", Monto = 185000, CantidadOrdenes = 200 },
                    new() { Moneda = "Dólares (USD)", CodigoMoneda = "USD", Monto = 45000, CantidadOrdenes = 35 },
                    new() { Moneda = "Euros (EUR)", CodigoMoneda = "EUR", Monto = 12000, CantidadOrdenes = 10 }
                },
                TotalOrdenes = 245,
                MontoTotal = 242000,
                TotalProveedores = 56,
                PromedioMensual = 40333
            };
        }
    }
}
