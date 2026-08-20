using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Data.SqlClient;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class EstadisticaPedidosController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<EstadisticaPedidosController> _logger;

        public EstadisticaPedidosController(DataContext context, ILogger<EstadisticaPedidosController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene los datos de estadística de pedidos de venta
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<EstadisticaPedidosDTO>> Get(
            [FromQuery] DateTime? fechaDesde,
            [FromQuery] DateTime? fechaHasta,
            [FromQuery] int? nroClie,
            [FromQuery] int? nroPed,
            [FromQuery] string? codExt,
            [FromQuery] int? status)
        {
            try
            {
                _logger.LogInformation($"Obteniendo estadística de pedidos: desde {fechaDesde?.ToString("yyyy-MM-dd")} hasta {fechaHasta?.ToString("yyyy-MM-dd")}, cliente {nroClie}, status {status}");

                var resultado = new EstadisticaPedidosDTO();

                // Parámetros para la consulta SQL
                var fechaDesdeParam = fechaDesde.HasValue ? new SqlParameter("@FechaDesde", fechaDesde.Value) : new SqlParameter("@FechaDesde", DBNull.Value);
                var fechaHastaParam = fechaHasta.HasValue ? new SqlParameter("@FechaHasta", fechaHasta.Value) : new SqlParameter("@FechaHasta", DBNull.Value);
                var nroClieParam = nroClie.HasValue ? new SqlParameter("@NroClie", nroClie.Value) : new SqlParameter("@NroClie", DBNull.Value);
                var nroPedParam = nroPed.HasValue ? new SqlParameter("@NroPed", nroPed.Value) : new SqlParameter("@NroPed", DBNull.Value);
                var codExtParam = !string.IsNullOrEmpty(codExt) ? new SqlParameter("@CodExt", codExt) : new SqlParameter("@CodExt", DBNull.Value);
                var statusParam = status.HasValue ? new SqlParameter("@Status", status.Value) : new SqlParameter("@Status", DBNull.Value);

                // Consulta SQL principal que une PEDENCA, PEDDETA y MOVISTO
                // Usamos una consulta dinámica para obtener los datos calculados
                var connection = _context.Database.GetDbConnection();
                await connection.OpenAsync();

                var query = @"
                    SELECT
                        P.NroPed,
                        P.FechEmis,
                        P.NroClie,
                        P.RasoClie,
                        D.CodiInt,
                        D.CodiExt,
                        D.Descrip,
                        D.CanPed,
                        ISNULL(D.CantEnt, 0) AS CantidadEntregada,
                        CASE
                            WHEN D.CanPed - ISNULL(D.CantEnt, 0) < 0 THEN 0
                            ELSE D.CanPed - ISNULL(D.CantEnt, 0)
                        END AS CantidadPendiente,
                        D.FeSolEnt,
                        CASE
                            WHEN ISNULL(D.CantEnt, 0) = 0 THEN 'PENDIENTE'
                            WHEN ISNULL(D.CantEnt, 0) < D.CanPed THEN 'PARCIAL'
                            ELSE 'CUMPLIDO'
                        END AS Estado
                    FROM PEDENCA P WITH (NOLOCK)
                    LEFT JOIN PEDDETA D WITH (NOLOCK) ON P.NroPed = D.NroPed AND D.Status >= 0 AND D.Status < 8
                    WHERE P.Status >= 0 AND P.Status < 8
                    AND D.NroPed IS NOT NULL
                    AND (@FechaDesde IS NULL OR P.FechEmis >= @FechaDesde)
                    AND (@FechaHasta IS NULL OR P.FechEmis <= @FechaHasta)
                    AND (@NroClie IS NULL OR P.NroClie = @NroClie)
                    AND (@NroPed IS NULL OR P.NroPed = @NroPed)
                    AND (@CodExt IS NULL OR D.CodiExt LIKE '%' + @CodExt + '%')
                    AND (@Status IS NULL OR P.Status = @Status)
                    ORDER BY P.NroPed, D.FeSolEnt";

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = query;

                    if (fechaDesde.HasValue)
                        command.Parameters.Add(new SqlParameter("@FechaDesde", fechaDesde.Value));
                    else
                        command.Parameters.Add(new SqlParameter("@FechaDesde", DBNull.Value));

                    if (fechaHasta.HasValue)
                        command.Parameters.Add(new SqlParameter("@FechaHasta", fechaHasta.Value));
                    else
                        command.Parameters.Add(new SqlParameter("@FechaHasta", DBNull.Value));

                    if (nroClie.HasValue)
                        command.Parameters.Add(new SqlParameter("@NroClie", nroClie.Value));
                    else
                        command.Parameters.Add(new SqlParameter("@NroClie", DBNull.Value));

                    if (nroPed.HasValue)
                        command.Parameters.Add(new SqlParameter("@NroPed", nroPed.Value));
                    else
                        command.Parameters.Add(new SqlParameter("@NroPed", DBNull.Value));

                    if (!string.IsNullOrEmpty(codExt))
                        command.Parameters.Add(new SqlParameter("@CodExt", codExt));
                    else
                        command.Parameters.Add(new SqlParameter("@CodExt", DBNull.Value));

                    if (status.HasValue)
                        command.Parameters.Add(new SqlParameter("@Status", status.Value));
                    else
                        command.Parameters.Add(new SqlParameter("@Status", DBNull.Value));

                    var detalles = new List<DetallePedidoEstadisticaDTO>();

                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            var detalle = new DetallePedidoEstadisticaDTO
                            {
                                NroPed = reader.GetInt32(0),
                                FechEmis = reader.IsDBNull(1) ? (DateTime?)null : reader.GetDateTime(1),
                                NroClie = reader.IsDBNull(2) ? (int?)null : reader.GetInt32(2),
                                RasoClie = reader.IsDBNull(3) ? null : reader.GetString(3),
                                CodiInt = reader.IsDBNull(4) ? (int?)null : reader.GetInt32(4),
                                CodiExt = reader.IsDBNull(5) ? null : reader.GetString(5),
                                Descrip = reader.IsDBNull(6) ? null : reader.GetString(6),
                                CanPed = reader.GetDecimal(7),
                                CantidadEntregada = reader.GetDecimal(8),
                                CantidadPendiente = reader.GetDecimal(9),
                                FeSolEnt = reader.IsDBNull(10) ? (DateTime?)null : reader.GetDateTime(10),
                                Estado = reader.GetString(11)
                            };
                            detalles.Add(detalle);
                        }
                    }

                    _logger.LogInformation($"Consulta SQL devolvió {detalles.Count} registros");

                    // Calcular indicadores principales
                    resultado.TotalPedidos = detalles.Select(d => d.NroPed).Distinct().Count();
                    resultado.TotalRenglones = detalles.Count;
                    resultado.CantidadSolicitada = detalles.Sum(d => d.CanPed);
                    resultado.CantidadRemitida = detalles.Sum(d => d.CantidadEntregada);
                    resultado.CantidadPendiente = detalles.Sum(d => d.CantidadPendiente);
                    resultado.PedidosAtrasados = 0;

                    // Evolución de pedidos por mes
                    resultado.EvolucionPedidos = detalles
                        .Where(d => d.FechEmis.HasValue)
                        .GroupBy(d => new { Año = d.FechEmis!.Value.Year, Mes = d.FechEmis!.Value.Month })
                        .Select(g => new DatoMensualPedidosDTO
                        {
                            Mes = $"{g.Key.Año}-{g.Key.Mes:D2}",
                            CantidadPedidos = g.Select(d => d.NroPed).Distinct().Count(),
                            CantidadSolicitada = g.Sum(d => d.CanPed)
                        })
                        .OrderBy(d => d.Mes)
                        .ToList();

                    // Pedidos por cliente (top 10)
                    resultado.PedidosPorCliente = detalles
                        .Where(d => d.NroClie.HasValue)
                        .GroupBy(d => d.NroClie!.Value)
                        .Select(g => new DatoClientePedidosDTO
                        {
                            NroClie = g.Key,
                            NombreCliente = g.First().RasoClie ?? $"Cliente {g.Key}",
                            CantidadPedidos = g.Select(d => d.NroPed).Distinct().Count(),
                            CantidadSolicitada = g.Sum(d => d.CanPed)
                        })
                        .OrderByDescending(c => c.CantidadSolicitada)
                        .Take(10)
                        .ToList();

                    // Estado de pedidos
                    resultado.EstadoPedidos = detalles
                        .GroupBy(d => d.Estado)
                        .Select(g => new DatoEstadoPedidosDTO
                        {
                            Estado = g.Key,
                            Cantidad = g.Count()
                        })
                        .ToList();

                    // Reporte detallado
                    resultado.ReporteDetallado = detalles;

                    _logger.LogInformation($"Estadística generada: {resultado.TotalPedidos} pedidos, {resultado.TotalRenglones} renglones, {resultado.PedidosAtrasados} atrasados");

                    return Ok(resultado);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener estadística de pedidos de venta");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el detalle de remitos para un pedido y producto específicos
        /// </summary>
        [HttpGet("remitos/{nroPed}/{codiInt}")]
        public async Task<ActionResult<List<DetalleRemitoEstadisticaDTO>>> ObtenerRemitosPorPedidoProducto(int nroPed, int codiInt)
        {
            try
            {
                _logger.LogInformation($"Obteniendo remitos para pedido {nroPed}, código {codiInt}");

                var connection = _context.Database.GetDbConnection();
                await connection.OpenAsync();

                var query = @"
                    SELECT
                        FECHMOV AS FechMov,
                        DOPRILAR AS Documento,
                        CANTSALID AS CantidadRemitida
                    FROM MOVISTO WITH (NOLOCK)
                    WHERE CODIMOVI = 'RT'
                    AND NUPEDCLI = @NroPed
                    AND COD_INTE = @CodiInt
                    ORDER BY FECHMOV";

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = query;
                    command.Parameters.Add(new SqlParameter("@NroPed", nroPed));
                    command.Parameters.Add(new SqlParameter("@CodiInt", codiInt));

                    var remitos = new List<DetalleRemitoEstadisticaDTO>();

                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            var remito = new DetalleRemitoEstadisticaDTO
                            {
                                FechMov = reader.GetDateTime(0),
                                Documento = reader.IsDBNull(1) ? null : reader.GetString(1),
                                CantidadRemitida = reader.GetDecimal(2)
                            };
                            remitos.Add(remito);
                        }
                    }

                    return Ok(remitos);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener remitos para pedido {NroPed}, código {CodiInt}", nroPed, codiInt);
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }
    }
}
