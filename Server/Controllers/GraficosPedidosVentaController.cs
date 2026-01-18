using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class GraficosPedidosVentaController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<GraficosPedidosVentaController> _logger;

        public GraficosPedidosVentaController(DataContext context, ILogger<GraficosPedidosVentaController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene todos los datos para los gráficos de pedidos de venta
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<GraficosPedidosVentaDTO>> Get([FromQuery] DateTime fechaDesde, [FromQuery] DateTime fechaHasta)
        {
            try
            {
                _logger.LogInformation($"Obteniendo datos de gráficos desde {fechaDesde:yyyy-MM-dd} hasta {fechaHasta:yyyy-MM-dd}");

                var resultado = new GraficosPedidosVentaDTO();

                // Obtener todos los pedidos del período
                var pedidos = await _context.Set<Pedenca>()
                    .Where(p => p.FechEmis >= fechaDesde && p.FechEmis <= fechaHasta)
                    .ToListAsync();

                // Obtener detalles de pedidos
                var nrosPedidos = pedidos.Select(p => p.NroPed).ToList();
                var detalles = await _context.Set<Peddeta>()
                    .Where(d => nrosPedidos.Contains(d.NroPed))
                    .ToListAsync();

                // Calcular totales por pedido desde los detalles
                // Si MoneEmis > 1, multiplicar por TipoCam para convertir a moneda base
                var totalesPorPedido = detalles
                    .Where(d => d.NroPed > 0)
                    .GroupBy(d => d.NroPed)
                    .ToDictionary(g => g.Key, g => g.Sum(d => 
                    {
                        var importe = (decimal)(d.ImpoTot ?? 0);
                        var tipoCambio = (decimal)(d.TipoCam ?? 1);
                        var moneEmis = d.MoneEmis ?? 1;
                        
                        // Si la moneda es diferente a la base (> 1), convertir multiplicando por tipo de cambio
                        return moneEmis > 1 ? importe * tipoCambio : importe;
                    }));

                // 1. VENTAS POR MES
                resultado.VentasPorMes = pedidos
                    .Where(p => p.FechEmis.HasValue)
                    .GroupBy(p => new { 
                        Año = p.FechEmis!.Value.Year, 
                        Mes = p.FechEmis!.Value.Month 
                    })
                    .Select(g => new DatoMensualDTO
                    {
                        Mes = $"{g.Key.Año}-{g.Key.Mes:D2}",
                        Monto = g.Sum(p => totalesPorPedido.ContainsKey(p.NroPed) ? totalesPorPedido[p.NroPed] : 0),
                        Cantidad = g.Count()
                    })
                    .OrderBy(d => d.Mes)
                    .ToList();

                // 2. TOP 10 CLIENTES
                var clientesIds = pedidos.Where(p => p.NroClie.HasValue).Select(p => p.NroClie!.Value).Distinct().ToList();
                var clientes = await _context.Set<Deudor12>()
                    .Where(c => clientesIds.Contains(c.Nume_Cli))
                    .ToDictionaryAsync(c => c.Nume_Cli, c => c.Raso_Cli ?? "Sin nombre");

                resultado.TopClientes = pedidos
                    .Where(p => p.NroClie.HasValue)
                    .GroupBy(p => p.NroClie!.Value)
                    .Select(g => new DatoClienteDTO
                    {
                        Nombre = clientes.ContainsKey(g.Key) ? clientes[g.Key] : $"Cliente {g.Key}",
                        Monto = g.Sum(p => totalesPorPedido.ContainsKey(p.NroPed) ? totalesPorPedido[p.NroPed] : 0),
                        Cantidad = g.Count()
                    })
                    .OrderByDescending(c => c.Monto)
                    .Take(10)
                    .ToList();

                // 3. ESTADOS DE PEDIDOS
                resultado.EstadosPedidos = pedidos
                    .GroupBy(p => p.Status ?? 0)
                    .Select(g => new DatoEstadoDTO
                    {
                        Estado = ObtenerNombreEstado(g.Key),
                        Cantidad = g.Count()
                    })
                    .ToList();

                // 4. MONTOS POR MONEDA
                var monedasIds = pedidos.Select(p => p.MoneEmis ?? 1).Distinct().ToList();
                var monedas = await _context.Set<Tamoned>()
                    .Where(m => monedasIds.Contains((short)m.CodigoMoneda))
                    .ToDictionaryAsync(m => (short)m.CodigoMoneda, m => m.Descripcion ?? "Sin nombre");

                resultado.MontosPorMoneda = pedidos
                    .GroupBy(p => p.MoneEmis ?? 1)
                    .Select(g => new DatoMonedaDTO
                    {
                        Moneda = monedas.ContainsKey(g.Key) ? monedas[g.Key] : $"Moneda {g.Key}",
                        Monto = g.Sum(p => totalesPorPedido.ContainsKey(p.NroPed) ? totalesPorPedido[p.NroPed] : 0)
                    })
                    .ToList();

                // 5. TOP 10 VENDEDORES
                var vendedoresIds = pedidos.Where(p => p.Vended.HasValue && p.Vended > 0).Select(p => p.Vended!.Value).Distinct().ToList();
                var vendedores = await _context.Set<Vendedor>()
                    .Where(v => v.NumVendedor.HasValue && vendedoresIds.Contains((short)v.NumVendedor.Value))
                    .ToDictionaryAsync(v => (short)v.NumVendedor!.Value, v => v.NombreVendedor ?? "Sin nombre");

                resultado.TopVendedores = pedidos
                    .Where(p => p.Vended.HasValue && p.Vended > 0)
                    .GroupBy(p => p.Vended!.Value)
                    .Select(g => new DatoVendedorDTO
                    {
                        Nombre = vendedores.ContainsKey(g.Key) ? vendedores[g.Key] : $"Vendedor {g.Key}",
                        Monto = g.Sum(p => totalesPorPedido.ContainsKey(p.NroPed) ? totalesPorPedido[p.NroPed] : 0),
                        Cantidad = g.Count()
                    })
                    .OrderByDescending(v => v.Monto)
                    .Take(10)
                    .ToList();

                // 6. TOP 10 PRODUCTOS MÁS VENDIDOS
                var productosIds = detalles.Where(d => d.CodiInt.HasValue).Select(d => d.CodiInt!.Value).Distinct().ToList();
                var productos = await _context.Set<Master>()
                    .Where(m => productosIds.Contains(m.Codint))
                    .ToDictionaryAsync(m => m.Codint, m => new { m.Codigo, m.Descripcion });

                resultado.TopProductos = detalles
                    .Where(d => d.CodiInt.HasValue)
                    .GroupBy(d => d.CodiInt!.Value)
                    .Select(g => new DatoProductoDTO
                    {
                        Codigo = productos.ContainsKey(g.Key) ? productos[g.Key].Codigo ?? "" : $"{g.Key}",
                        Descripcion = productos.ContainsKey(g.Key) ? productos[g.Key].Descripcion ?? "" : $"Producto {g.Key}",
                        Cantidad = g.Sum(d => (decimal)(d.CanPed ?? 0)),
                        Monto = g.Sum(d => 
                        {
                            var importe = (decimal)(d.ImpoTot ?? 0);
                            var tipoCambio = (decimal)(d.TipoCam ?? 1);
                            var moneEmis = d.MoneEmis ?? 1;
                            return moneEmis > 1 ? importe * tipoCambio : importe;
                        })
                    })
                    .OrderByDescending(p => p.Cantidad)
                    .Take(10)
                    .ToList();

                // ESTADÍSTICAS GENERALES
                resultado.TotalPedidos = pedidos.Count();
                resultado.MontoTotal = totalesPorPedido.Values.Sum();
                resultado.TotalClientes = pedidos.Select(p => p.NroClie).Distinct().Count();
                resultado.TotalVendedores = pedidos.Where(p => p.Vended.HasValue && p.Vended > 0).Select(p => p.Vended).Distinct().Count();
                resultado.TotalProductos = detalles.Select(d => d.CodiInt).Distinct().Count();
                
                // Calcular promedio mensual
                var mesesConDatos = resultado.VentasPorMes.Count;
                resultado.PromedioMensual = mesesConDatos > 0 ? resultado.MontoTotal / mesesConDatos : 0;

                _logger.LogInformation($"Datos generados: {resultado.TotalPedidos} pedidos, {resultado.TotalClientes} clientes");

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener datos de gráficos de pedidos de venta");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el nombre descriptivo del estado
        /// </summary>
        private string ObtenerNombreEstado(short status)
        {
            return status switch
            {
                -1 => "Pendiente",
                0 => "Activo",
                9 => "Anulado",
                _ => $"Estado {status}"
            };
        }
    }
}
