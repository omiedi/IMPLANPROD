using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Shared.DTOs;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/reportes")]
    public class ReportesComprasController : ControllerBase
    {
        private readonly DataContext _context;

        public ReportesComprasController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene estadísticas resumen de compras en un período
        /// </summary>
        [HttpGet("estadisticas-resumen")]
        public async Task<ActionResult<EstadisticasResumenDTO>> GetEstadisticasResumen(DateTime fechaDesde, DateTime fechaHasta)
        {
            try
            {
                var query = _context.Ocomencas
                    .Where(o => o.FGenOcom.HasValue && 
                               o.FGenOcom >= fechaDesde && 
                               o.FGenOcom <= fechaHasta);

                var totalOrdenes = await query.CountAsync();
                var montoTotal = await query.SumAsync(o => o.TotalOcom ?? 0);
                var totalProveedores = await query
                    .Where(o => o.NproOcom.HasValue)
                    .Select(o => o.NproOcom)
                    .Distinct()
                    .CountAsync();

                var mesesEnPeriodo = ((fechaHasta.Year - fechaDesde.Year) * 12) + fechaHasta.Month - fechaDesde.Month + 1;
                var promedioMensual = mesesEnPeriodo > 0 ? montoTotal / mesesEnPeriodo : 0;

                return Ok(new
                {
                    TotalOrdenes = totalOrdenes,
                    MontoTotal = montoTotal,
                    TotalProveedores = totalProveedores,
                    PromedioMensual = promedioMensual
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en estadísticas resumen: {ex.Message}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene datos de compras agrupadas por mes
        /// </summary>
        [HttpGet("compras-por-mes")]
        public async Task<ActionResult<object>> GetComprasPorMes(DateTime fechaDesde, DateTime fechaHasta)
        {
            try
            {
                var datos = await _context.Ocomencas
                    .Where(o => o.FGenOcom.HasValue && 
                               o.FGenOcom >= fechaDesde && 
                               o.FGenOcom <= fechaHasta)
                    .GroupBy(o => new { 
                        Año = o.FGenOcom.Value.Year, 
                        Mes = o.FGenOcom.Value.Month 
                    })
                    .Select(g => new
                    {
                        Mes = $"{g.Key.Año}-{g.Key.Mes:00}",
                        Total = g.Sum(o => o.TotalOcom ?? 0),
                        Cantidad = g.Count()
                    })
                    .OrderBy(x => x.Mes)
                    .ToListAsync();

                return Ok(datos);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en compras por mes: {ex.Message}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene top proveedores por monto
        /// </summary>
        [HttpGet("top-proveedores")]
        public async Task<ActionResult<object>> GetTopProveedores(DateTime fechaDesde, DateTime fechaHasta, int limite = 10)
        {
            try
            {
                var datos = await _context.Ocomencas
                    .Where(o => o.FGenOcom.HasValue && 
                               o.FGenOcom >= fechaDesde && 
                               o.FGenOcom <= fechaHasta &&
                               !string.IsNullOrEmpty(o.RaSoProv))
                    .GroupBy(o => o.RaSoProv)
                    .Select(g => new
                    {
                        Proveedor = g.Key.Length > 20 ? g.Key.Substring(0, 20) + "..." : g.Key,
                        ProveedorCompleto = g.Key,
                        Total = g.Sum(o => o.TotalOcom ?? 0),
                        Cantidad = g.Count()
                    })
                    .OrderByDescending(x => x.Total)
                    .Take(limite)
                    .ToListAsync();

                return Ok(datos);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en top proveedores: {ex.Message}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene distribución de órdenes por estado
        /// </summary>
        [HttpGet("estados-ordenes")]
        public async Task<ActionResult<object>> GetEstadosOrdenes(DateTime fechaDesde, DateTime fechaHasta)
        {
            try
            {
                // Como no tenemos campo de estado específico, usamos lógica basada en fechas
                var ordenes = await _context.Ocomencas
                    .Where(o => o.FGenOcom.HasValue && 
                               o.FGenOcom >= fechaDesde && 
                               o.FGenOcom <= fechaHasta)
                    .Select(o => new
                    {
                        o.FGenOcom,
                        o.FAproOcom,
                        o.FAnuOcom,
                        o.TotalOcom
                    })
                    .ToListAsync();

                var estados = ordenes
                    .GroupBy(o => 
                        o.FAnuOcom.HasValue ? "Anulada" :
                        o.FAproOcom.HasValue ? "Aprobada" :
                        "Pendiente")
                    .Select(g => new
                    {
                        Estado = g.Key,
                        Cantidad = g.Count(),
                        MontoTotal = g.Sum(o => o.TotalOcom ?? 0)
                    })
                    .ToList();

                return Ok(estados);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en estados órdenes: {ex.Message}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene distribución de montos por moneda
        /// </summary>
        [HttpGet("montos-por-moneda")]
        public async Task<ActionResult<object>> GetMontosPorMoneda(DateTime fechaDesde, DateTime fechaHasta)
        {
            try
            {
                var datos = await _context.Ocomencas
                    .Where(o => o.FGenOcom.HasValue && 
                               o.FGenOcom >= fechaDesde && 
                               o.FGenOcom <= fechaHasta &&
                               o.TotalOcom.HasValue && o.TotalOcom > 0)
                    .GroupBy(o => o.MoneEmis ?? 1)
                    .Select(g => new
                    {
                        CodigoMoneda = g.Key,
                        Moneda = g.Key == 1 ? "Pesos" : 
                                g.Key == 2 ? "Dólares" : 
                                $"Moneda {g.Key}",
                        Total = g.Sum(o => o.TotalOcom ?? 0),
                        Cantidad = g.Count()
                    })
                    .OrderByDescending(x => x.Total)
                    .ToListAsync();

                return Ok(datos);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en montos por moneda: {ex.Message}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

    }
}
