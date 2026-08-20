using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class GraficosRemitosController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<GraficosRemitosController> _logger;

        public GraficosRemitosController(DataContext context, ILogger<GraficosRemitosController> logger)
        {
            _context = context;
            _logger = logger;
        }

        [HttpGet]
        public async Task<ActionResult<GraficosRemitosDTO>> GetGraficosRemitos(
            [FromQuery] DateTime? fechaDesde = null,
            [FromQuery] DateTime? fechaHasta = null)
        {
            try
            {
                var desde = fechaDesde ?? DateTime.Now.AddMonths(-6);
                var hasta = fechaHasta ?? DateTime.Now;

                _logger.LogInformation("Obteniendo gráficos de remitos desde {Desde} hasta {Hasta}", desde, hasta);

                var resultado = new GraficosRemitosDTO
                {
                    FechaDesde = desde,
                    FechaHasta = hasta
                };

                // Código de movimiento configurado para remitos (CODREFAC) o RT por defecto
                var codigoMov = await ObtenerCodigoMovimientoConfiguradoAsync();
                if (string.IsNullOrWhiteSpace(codigoMov) || codigoMov.Length != 2)
                {
                    codigoMov = "RT";
                }

                var remitosQuery = _context.Set<Movisto>()
                    .AsNoTracking()
                    .Where(m => m.FechMov >= desde &&
                                m.FechMov <= hasta &&
                                (m.CodiMovi == codigoMov || m.CodiMovi == "DV") &&
                                m.CantSalid > 0);

                var remitos = await remitosQuery.ToListAsync();

                if (!remitos.Any())
                {
                    _logger.LogWarning("No se encontraron remitos en el rango especificado");
                    return Ok(GenerarDatosEjemplo(desde, hasta));
                }

                resultado.RemitosPorMes = await GenerarRemitosPorMes(remitos, desde, hasta);
                resultado.TopClientes = await GenerarTopClientes(remitos);
                resultado.EstadosRemitos = await GenerarEstadosRemitos(remitos);
                resultado.RemitosPorDeposito = await GenerarRemitosPorDeposito(remitos);
                resultado.TopProductos = await GenerarTopProductos(remitos);

                resultado.TotalRemitos = remitos
                    .Where(r => r.DoPriCor != null)
                    .GroupBy(r => new { r.DoPriCor, r.Nume_Clie, r.FechMov })
                    .Count();

                // En la mayoría de los remitos no hay importe informado, por lo que
                // no se utiliza el monto como métrica principal en los gráficos.
                // Se deja MontoTotal en cero para evitar interpretaciones erróneas.
                resultado.MontoTotal = 0;
                resultado.TotalClientes = remitos
                    .Where(r => r.Nume_Clie.HasValue)
                    .Select(r => r.Nume_Clie!.Value)
                    .Distinct()
                    .Count();

                // Total de productos distintos que tuvieron salida en el período
                resultado.TotalProductosDistintos = remitos
                    .Where(r => r.Cod_Inte.HasValue && r.CantSalid.HasValue && r.CantSalid.Value > 0)
                    .Select(r => r.Cod_Inte!.Value)
                    .Distinct()
                    .Count();

                var mesesEnRango = ((hasta.Year - desde.Year) * 12) + hasta.Month - desde.Month + 1;
                // PromedioMensual ahora representa promedio de remitos por mes
                resultado.PromedioMensual = mesesEnRango > 0
                    ? (decimal)resultado.TotalRemitos / mesesEnRango
                    : 0;

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
                //_logger.LogError(ex, "Error al obtener datos para gráficos de remitos");
                //return StatusCode(500, "Error interno del servidor");
            }
        }

        private async Task<string> ObtenerCodigoMovimientoConfiguradoAsync()
        {
            try
            {
                var config = await _context.Set<FLEXCRL>()
                    .AsNoTracking()
                    .FirstOrDefaultAsync(f => f.CLABUS == "" && f.PROBUS == "CODREFAC");

                return config?.VALCONTROL ?? "RT";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener código de movimiento configurado para remitos");
                return "RT";
            }
        }

        private async Task<List<DatoMensualRemitoDTO>> GenerarRemitosPorMes(List<Movisto> remitos, DateTime desde, DateTime hasta)
        {
            try
            {
                var datosPorMes = remitos
                    .Where(r => r.FechMov.HasValue)
                    .GroupBy(r => new
                    {
                        Año = r.FechMov!.Value.Year,
                        Mes = r.FechMov!.Value.Month
                    })
                    .Select(g => new DatoMensualRemitoDTO
                    {
                        // Usar formato yyyy-MM para compatibilidad con ApexCharts
                        Mes = $"{g.Key.Año:D4}-{g.Key.Mes:D2}",
                        // Monto queda disponible por compatibilidad, pero los gráficos
                        // utilizarán Cantidad (cantidad de remitos) como métrica.
                        Monto = g.Sum(r => r.ValTotItIva ?? r.ValTotIt ?? 0),
                        Cantidad = g
                            .Where(r => r.DoPriCor != null)
                            .GroupBy(r => new { r.DoPriCor, r.Nume_Clie, r.FechMov })
                            .Count(),
                        Año = g.Key.Año,
                        NumeroMes = g.Key.Mes
                    })
                    .OrderBy(d => d.Año)
                    .ThenBy(d => d.NumeroMes)
                    .ToList();

                return await Task.FromResult(datosPorMes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar datos de remitos por mes");
                return new List<DatoMensualRemitoDTO>();
            }
        }

        private async Task<List<DatoClienteRemitoDTO>> GenerarTopClientes(List<Movisto> remitos)
        {
            try
            {
                var grupos = remitos
                    .Where(r => r.Nume_Clie.HasValue)
                    .GroupBy(r => r.Nume_Clie!.Value)
                    .ToList();

                var numerosClientes = grupos.Select(g => g.Key).ToList();

                var clientes = await _context.Set<Deudor12>()
                    .Where(d => numerosClientes.Contains(d.Nume_Cli))
                    .Select(d => new { d.Nume_Cli, d.Raso_Cli })
                    .ToListAsync();

                var dicClientes = clientes.ToDictionary(
                    c => c.Nume_Cli,
                    c => c.Raso_Cli ?? $"CLIENTE {c.Nume_Cli}");

                var topClientes = grupos
                    .Select(g => new DatoClienteRemitoDTO
                    {
                        Nombre = dicClientes.ContainsKey(g.Key)
                            ? TruncarTexto(dicClientes[g.Key], 30)
                            : $"CLIENTE {g.Key}",
                        // Monto se mantiene por compatibilidad, pero el gráfico usará CantidadRemitos.
                        Monto = g.Sum(r => r.ValTotItIva ?? r.ValTotIt ?? 0),
                        CantidadRemitos = g
                            .Where(r => r.DoPriCor != null)
                            .GroupBy(r => new { r.DoPriCor, r.Nume_Clie, r.FechMov })
                            .Count()
                    })
                    .OrderByDescending(c => c.CantidadRemitos)
                    .Take(5)
                    .ToList();

                return topClientes;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar top clientes de remitos");
                return new List<DatoClienteRemitoDTO>();
            }
        }

        private async Task<List<DatoEstadoRemitoDTO>> GenerarEstadosRemitos(List<Movisto> remitos)
        {
            try
            {
                var estados = remitos
                    .GroupBy(r => ObtenerEstadoRemito(r))
                    .Select(g => new DatoEstadoRemitoDTO
                    {
                        Estado = g.Key,
                        Cantidad = g
                            .Where(r => r.DoPriCor != null)
                            .GroupBy(r => new { r.DoPriCor, r.Nume_Clie, r.FechMov })
                            .Count(),
                        // MontoTotal no se utiliza actualmente en los gráficos, se mantiene
                        // solo por compatibilidad.
                        MontoTotal = g.Sum(r => r.ValTotItIva ?? r.ValTotIt ?? 0)
                    })
                    .OrderByDescending(e => e.Cantidad)
                    .ToList();

                return await Task.FromResult(estados);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar estados de remitos");
                return new List<DatoEstadoRemitoDTO>();
            }
        }

        private async Task<List<DatoDepositoRemitoDTO>> GenerarRemitosPorDeposito(List<Movisto> remitos)
        {
            try
            {
                var grupos = remitos
                    .Where(r => r.DepoMovi.HasValue && r.DepoMovi.Value > 0)
                    .GroupBy(r => r.DepoMovi!.Value)
                    .ToList();

                var codigosDeposito = grupos.Select(g => g.Key).ToList();

                var depositos = await _context.Set<Tadeposi>()
                    .Where(d => codigosDeposito.Contains(d.CodigoDepo))
                    .Select(d => new { d.CodigoDepo, d.Descripcion })
                    .ToListAsync();

                var dicDepositos = depositos.ToDictionary(
                    d => d.CodigoDepo,
                    d => d.Descripcion ?? $"Depósito {d.CodigoDepo}");

                var datos = grupos
                    .Select(g => new DatoDepositoRemitoDTO
                    {
                        CodigoDeposito = g.Key,
                        Deposito = dicDepositos.ContainsKey(g.Key)
                            ? TruncarTexto(dicDepositos[g.Key], 30)
                            : $"Depósito {g.Key}",
                        // Monto se mantiene por compatibilidad, pero no se utilizará
                        // en los gráficos ya que se priorizan las cantidades.
                        Monto = g.Sum(r => r.ValTotItIva ?? r.ValTotIt ?? 0),
                        CantidadRemitos = g
                            .Where(r => r.DoPriCor != null)
                            .GroupBy(r => new { r.DoPriCor, r.Nume_Clie, r.FechMov })
                            .Count()
                    })
                    .OrderByDescending(d => d.CantidadRemitos)
                    .ToList();

                return datos;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar remitos por depósito");
                return new List<DatoDepositoRemitoDTO>();
            }
        }

        private async Task<List<DatoProductoRemitoDTO>> GenerarTopProductos(List<Movisto> remitos)
        {
            try
            {
                // Agrupar por código interno de producto para obtener cantidades y remitos
                var grupos = remitos
                    .Where(r => r.Cod_Inte.HasValue && r.CantSalid.HasValue && r.CantSalid.Value > 0)
                    .GroupBy(r => r.Cod_Inte!.Value)
                    .ToList();

                if (!grupos.Any())
                {
                    return new List<DatoProductoRemitoDTO>();
                }

                var codigos = grupos.Select(g => g.Key).ToList();

                // Obtener descripciones desde Master
                var productos = await _context.Set<Master>()
                    .Where(m => codigos.Contains(m.Codint))
                    .Select(m => new { m.Codint, m.Descripcion })
                    .ToListAsync();

                var dicProductos = productos.ToDictionary(
                    p => p.Codint,
                    p => p.Descripcion ?? $"Prod. {p.Codint}");

                var topProductos = grupos
                    .Select(g => new DatoProductoRemitoDTO
                    {
                        Producto = dicProductos.ContainsKey(g.Key)
                            ? TruncarTexto(dicProductos[g.Key], 30)
                            : $"Prod. {g.Key}",
                        CantidadUnidades = (int)g.Sum(r => r.CantSalid ?? 0),
                        CantidadRemitos = g
                            .Where(r => r.DoPriCor != null)
                            .GroupBy(r => new { r.DoPriCor, r.Nume_Clie, r.FechMov })
                            .Count()
                    })
                    .OrderByDescending(p => p.CantidadUnidades)
                    .Take(5)
                    .ToList();

                return topProductos;
            }
            catch (Exception ex)
            {

                _logger.LogError(ex, "Error al generar top productos de remitos");
                return new List<DatoProductoRemitoDTO>();
            }
        }

        private string ObtenerNombreMes(int mes)
        {
            var meses = new[] { "", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic" };
            return mes >= 1 && mes <= 12 ? meses[mes] : mes.ToString();
        }

        private string ObtenerEstadoRemito(Movisto remito)
        {
            if (!string.IsNullOrEmpty(remito.ReferCor) && remito.ReferCor.StartsWith("(Anul)"))
            {
                return "Anulado";
            }

            return "Vigente";
        }

        private string TruncarTexto(string texto, int maxLength)
        {
            if (string.IsNullOrEmpty(texto)) return string.Empty;
            return texto.Length <= maxLength ? texto : texto.Substring(0, maxLength) + "...";
        }

        private GraficosRemitosDTO GenerarDatosEjemplo(DateTime desde, DateTime hasta)
        {
            return new GraficosRemitosDTO
            {
                FechaDesde = desde,
                FechaHasta = hasta,
                RemitosPorMes = new List<DatoMensualRemitoDTO>
                {
                    new() { Mes = "Ene 2024", Monto = 120000, Cantidad = 80, Año = 2024, NumeroMes = 1 },
                    new() { Mes = "Feb 2024", Monto = 150000, Cantidad = 95, Año = 2024, NumeroMes = 2 },
                    new() { Mes = "Mar 2024", Monto = 110000, Cantidad = 70, Año = 2024, NumeroMes = 3 },
                },
                TopClientes = new List<DatoClienteRemitoDTO>
                {
                    new() { Nombre = "CLIENTE A", Monto = 75000, CantidadRemitos = 25 },
                    new() { Nombre = "CLIENTE B", Monto = 60000, CantidadRemitos = 20 },
                    new() { Nombre = "CLIENTE C", Monto = 45000, CantidadRemitos = 15 },
                },
                EstadosRemitos = new List<DatoEstadoRemitoDTO>
                {
                    new() { Estado = "Vigente", Cantidad = 210, MontoTotal = 320000 },
                    new() { Estado = "Anulado", Cantidad = 15, MontoTotal = 18000 },
                },
                RemitosPorDeposito = new List<DatoDepositoRemitoDTO>
                {
                    new() { CodigoDeposito = 1, Deposito = "Depósito Principal", Monto = 250000, CantidadRemitos = 150 },
                    new() { CodigoDeposito = 2, Deposito = "Depósito Secundario", Monto = 90000, CantidadRemitos = 60 },
                },
                TotalRemitos = 225,
                MontoTotal = 338000,
                TotalClientes = 35,
                PromedioMensual = 112666
            };
        }
    }
}
