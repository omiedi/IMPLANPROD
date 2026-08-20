using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Text.RegularExpressions;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PeddetaController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<PeddetaController> _logger;

        public PeddetaController(DataContext context, ILogger<PeddetaController> logger)
        {
            _context = context;
            _logger = logger;
        }

        [HttpGet("pendientes-detalle")]
        public async Task<ActionResult<List<PedidoVentaPendienteDetalleDTO>>> ObtenerPedidosPendientesDetalle(
            [FromQuery] string? filter = null,
            [FromQuery] int? nroPed = null,
            [FromQuery] string? codigo = null,
            [FromQuery] string? descripcion = null,
            [FromQuery] string? uMedida = null,
            [FromQuery] DateTime? fechaEmisDesde = null,
            [FromQuery] DateTime? fechaEmisHasta = null,
            [FromQuery] decimal? cantSolicitadaMin = null,
            [FromQuery] decimal? cantSolicitadaMax = null,
            [FromQuery] decimal? cantEntregadaMin = null,
            [FromQuery] decimal? cantEntregadaMax = null,
            [FromQuery] decimal? saldoMin = null,
            [FromQuery] decimal? saldoMax = null,
            [FromQuery] string? orderBy = null,
            [FromQuery] bool asc = true)
        {
            try
            {
                var q =
                    from pd in _context.Peddetas.AsNoTracking()
                    group pd by new
                    {
                        pd.NroPed,
                        pd.CodiInt,
                        pd.CodiExt,
                        pd.Descrip
                    }
                    into g
                    select new
                    {
                        g.Key,
                        CantPed = g.Sum(x => x.CanPed ?? 0m),
                        CantEnt = g.Sum(x => x.CantEnt ?? 0m),
                        FechaEntrega = g.Min(x => x.FeSolEnt)
                    }
                    into x
                    where (x.CantPed - x.CantEnt) > 0.005m
                    join pe in _context.Pedencas.AsNoTracking() on (x.Key.NroPed ?? 0) equals (pe.NroPed ?? 0) into pedencaJoin
                    from pe in pedencaJoin.DefaultIfEmpty()
                    join m in _context.masters.AsNoTracking()
                        on (x.Key.CodiInt ?? 0) equals m.Codint into masterJoin
                    from m in masterJoin.DefaultIfEmpty()
                    select new
                    {
                        Item = new PedidoVentaPendienteDetalleDTO
                        {
                            NroPed = x.Key.NroPed ?? 0,
                            Codigo = x.Key.CodiExt ?? "",
                            Descripcion = x.Key.Descrip ?? "",
                            CantSolicitada = x.CantPed,
                            CantEntregada = x.CantEnt,
                            Saldo = x.CantPed - x.CantEnt,
                            UMedida = m != null ? (m.Umedida ?? "") : "",
                            FechaEntrega = x.FechaEntrega
                        },
                        FechaEmision = pe != null ? pe.FechEmis : null
                    };

                if (fechaEmisDesde.HasValue)
                {
                    var fd = fechaEmisDesde.Value.Date;
                    q = q.Where(x => x.FechaEmision.HasValue && x.FechaEmision.Value.Date >= fd);
                }

                if (fechaEmisHasta.HasValue)
                {
                    var fh = fechaEmisHasta.Value.Date;
                    q = q.Where(x => x.FechaEmision.HasValue && x.FechaEmision.Value.Date <= fh);
                }

                var qItems = q.Select(x => x.Item);

                if (nroPed.HasValue)
                {
                    qItems = qItems.Where(x => x.NroPed == nroPed.Value);
                }

                if (!string.IsNullOrWhiteSpace(codigo))
                {
                    var c = codigo.Trim().ToLower();
                    qItems = qItems.Where(x => x.Codigo != null && x.Codigo.ToLower().Contains(c));
                }

                if (!string.IsNullOrWhiteSpace(descripcion))
                {
                    var d = descripcion.Trim().ToLower();
                    qItems = qItems.Where(x => x.Descripcion != null && x.Descripcion.ToLower().Contains(d));
                }

                if (!string.IsNullOrWhiteSpace(uMedida))
                {
                    var u = uMedida.Trim().ToLower();
                    qItems = qItems.Where(x => x.UMedida != null && x.UMedida.ToLower().Contains(u));
                }

                if (cantSolicitadaMin.HasValue)
                {
                    qItems = qItems.Where(x => x.CantSolicitada >= cantSolicitadaMin.Value);
                }

                if (cantSolicitadaMax.HasValue)
                {
                    qItems = qItems.Where(x => x.CantSolicitada <= cantSolicitadaMax.Value);
                }

                if (cantEntregadaMin.HasValue)
                {
                    qItems = qItems.Where(x => x.CantEntregada >= cantEntregadaMin.Value);
                }

                if (cantEntregadaMax.HasValue)
                {
                    qItems = qItems.Where(x => x.CantEntregada <= cantEntregadaMax.Value);
                }

                if (saldoMin.HasValue)
                {
                    qItems = qItems.Where(x => x.Saldo >= saldoMin.Value);
                }

                if (saldoMax.HasValue)
                {
                    qItems = qItems.Where(x => x.Saldo <= saldoMax.Value);
                }

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    var f = filter.Trim().ToLower();
                    qItems = qItems.Where(x =>
                        x.NroPed.ToString().Contains(f) ||
                        (x.Codigo != null && x.Codigo.ToLower().Contains(f)) ||
                        (x.Descripcion != null && x.Descripcion.ToLower().Contains(f)) ||
                        (x.UMedida != null && x.UMedida.ToLower().Contains(f)) ||
                        (x.FechaEntrega.HasValue && x.FechaEntrega.Value.ToString("dd/MM/yyyy").Contains(f)) ||
                        x.CantSolicitada.ToString().Contains(f) ||
                        x.CantEntregada.ToString().Contains(f) ||
                        x.Saldo.ToString().Contains(f)
                    );
                }

                orderBy = (orderBy ?? string.Empty).Trim();
                qItems = orderBy switch
                {
                    nameof(PedidoVentaPendienteDetalleDTO.NroPed) => asc ? qItems.OrderBy(x => x.NroPed) : qItems.OrderByDescending(x => x.NroPed),
                    nameof(PedidoVentaPendienteDetalleDTO.Codigo) => asc ? qItems.OrderBy(x => x.Codigo) : qItems.OrderByDescending(x => x.Codigo),
                    nameof(PedidoVentaPendienteDetalleDTO.Descripcion) => asc ? qItems.OrderBy(x => x.Descripcion) : qItems.OrderByDescending(x => x.Descripcion),
                    nameof(PedidoVentaPendienteDetalleDTO.CantSolicitada) => asc ? qItems.OrderBy(x => x.CantSolicitada) : qItems.OrderByDescending(x => x.CantSolicitada),
                    nameof(PedidoVentaPendienteDetalleDTO.UMedida) => asc ? qItems.OrderBy(x => x.UMedida) : qItems.OrderByDescending(x => x.UMedida),
                    nameof(PedidoVentaPendienteDetalleDTO.CantEntregada) => asc ? qItems.OrderBy(x => x.CantEntregada) : qItems.OrderByDescending(x => x.CantEntregada),
                    nameof(PedidoVentaPendienteDetalleDTO.Saldo) => asc ? qItems.OrderBy(x => x.Saldo) : qItems.OrderByDescending(x => x.Saldo),
                    nameof(PedidoVentaPendienteDetalleDTO.FechaEntrega) => asc ? qItems.OrderBy(x => x.FechaEntrega) : qItems.OrderByDescending(x => x.FechaEntrega),
                    _ => qItems.OrderBy(x => x.NroPed).ThenBy(x => x.Codigo)
                };

                var rows = await qItems.ToListAsync();
                rows = rows
                    .Select(x =>
                    {
                        if (x.Saldo < 0m) x.Saldo = 0m;
                        return x;
                    })
                    .Where(x => x.Saldo > 0m)
                    .ToList();

                return Ok(rows);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener detalle de pedidos pendientes");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        [HttpGet("pendientes-por-producto/{codint:int}")]
        public async Task<ActionResult<List<StockPedidoPendienteDTO>>> ObtenerPedidosPendientesPorProducto(int codint)
        {
            try
            {
                var query = from pd in _context.Peddetas.AsNoTracking()
                            join c in _context.Clientes.AsNoTracking()
                                on (pd.NroClie ?? 0) equals c.Nume_Cli into clienteJoin
                            from c in clienteJoin.DefaultIfEmpty()
                            where pd.CodiInt == codint
                                  && pd.VenRep == "V"
                                  && (pd.Status ?? 0) < 8
                                  && (pd.CanPed ?? 0m) - (pd.CantEnt ?? 0m) > 0.005m
                            orderby pd.NroPed
                            select new StockPedidoPendienteDTO
                            {
                                NroPed = pd.NroPed ?? 0,
                                NroClie = pd.NroClie ?? 0,
                                Cliente = c != null ? (c.Raso_Cli ?? string.Empty) : string.Empty,
                                CantPendiente = (pd.CanPed ?? 0m) - (pd.CantEnt ?? 0m),
                                SalPreparado = 0m,
                                FechaSolEnt = pd.FeSolEnt
                            };

                var pedidos = await query.ToListAsync();
                pedidos = pedidos
                    .Select(x =>
                    {
                        if (x.CantPendiente < 0m)
                        {
                            x.CantPendiente = 0m;
                        }
                        return x;
                    })
                    .Where(x => x.CantPendiente > 0m)
                    .ToList();

                return Ok(pedidos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener pedidos pendientes por producto {Codint}", codint);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene los items de pedidos pendientes de entrega para un cliente (usado en /remitos/traslado)
        /// </summary>
        /// <param name="numeroCliente">Número del cliente</param>
        /// <returns>Lista de items pendientes con saldo mayor a 0</returns>
        [HttpGet("pendientes/{numeroCliente}")]
        public async Task<ActionResult<List<Peddeta>>> ObtenerPedidosPendientes(int numeroCliente)
        {
            try
            {
                _logger.LogInformation($"Obteniendo pedidos pendientes para cliente: {numeroCliente}");

                // Usar FromSqlRaw para asegurar que se cargue NuOrItem
                var pedidosPendientes = await _context.Set<Peddeta>()
                    .FromSqlRaw(@"
                        SELECT * FROM PEDDETA
                        WHERE NroClie = {0}
                        AND CanPed > ISNULL(CantEnt, 0)
                        AND Status >= 0
                        AND Status < 8
                        ORDER BY FeSolEnt", numeroCliente)
                    .ToListAsync();

                _logger.LogInformation($"Se encontraron {pedidosPendientes.Count} items pendientes para cliente {numeroCliente}");

                // Log para verificar NuOrItem
                foreach (var item in pedidosPendientes.Take(5))
                {
                    _logger.LogInformation($"Item: NroPed={item.NroPed}, CodiInt={item.CodiInt}, NuOrItem={item.NuOrItem}, CanSaldo={item.CanSaldo}");
                }

                return Ok(pedidosPendientes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener pedidos pendientes para cliente {numeroCliente}");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene los items de pedidos pendientes de entrega para un cliente (usado en /remitos/cliente)
        /// Considera transferencias y remitos para calcular pendiente de remitir
        /// </summary>
        /// <param name="numeroCliente">Número del cliente</param>
        /// <param name="nroPed">Número de pedido opcional para filtrar</param>
        /// <returns>Lista de items pendientes con información de transferencias y remitos</returns>
        [HttpGet("pendientes-cliente/{numeroCliente}")]
        public async Task<ActionResult<List<Peddeta>>> ObtenerPedidosPendientesCliente(int numeroCliente, [FromQuery] int? nroPed = null)
        {
            try
            {
                _logger.LogInformation($"Obteniendo pedidos pendientes para cliente: {numeroCliente}, pedido: {nroPed}");

                var nroPedParam = nroPed.HasValue ? new SqlParameter("@NroPed", nroPed.Value) : new SqlParameter("@NroPed", DBNull.Value);
                var nroClieParam = new SqlParameter("@NroClie", numeroCliente);

                var pedidosPendientes = await _context.Set<Peddeta>()
                    .FromSqlRaw(@"
                        SELECT
                            P.*,

                            -- Cantidad transferida para el pedido y código
                            ISNULL(M.CantidadTransferida, 0) AS CantidadTransferida,

                            -- Cantidad ya remitida para el pedido y código
                            ISNULL(M.CantidadRemitida, 0) AS CantidadRemitida,

                            -- Pendiente de remitir
                            CASE
                                WHEN
                                    ISNULL(M.CantidadTransferida, 0)
                                    - ISNULL(M.CantidadRemitida, 0) < 0
                                THEN 0
                                ELSE
                                    ISNULL(M.CantidadTransferida, 0)
                                    - ISNULL(M.CantidadRemitida, 0)
                            END AS CantidadPendiente

                        FROM PEDDETA P WITH (NOLOCK)

                        LEFT JOIN
                        (
                            SELECT
                                NUPEDCLI,
                                COD_INTE,

                                -- Total de transferencias
                                SUM(
                                    CASE
                                        WHEN CODIMOVI = 'TR'
                                        THEN CANTSALID
                                        ELSE 0
                                    END
                                ) AS CantidadTransferida,

                                -- Total de remitos
                                SUM(
                                    CASE
                                        WHEN CODIMOVI = 'RT'
                                        THEN CANTSALID
                                        ELSE 0
                                    END
                                ) AS CantidadRemitida

                            FROM MOVISTO WITH (NOLOCK)

                            WHERE CODIMOVI IN ('TR', 'RT')

                            GROUP BY
                                NUPEDCLI,
                                COD_INTE

                        ) M
                            ON M.NUPEDCLI = P.NroPed
                            AND M.COD_INTE = P.CodiInt

                        WHERE P.NroClie = @NroClie
                          AND P.Status < 8
                          AND P.Status >= 0

                          AND (@NroPed IS NULL OR P.NroPed = @NroPed)

                          AND
                          (
                              (P.CanPed - P.CantEnt) >= 0.05

                              OR

                              (
                                  P.TipoPed = 'ABIERTO'
                                  AND P.FinVigen >= GETDATE()
                              )
                          )

                        ORDER BY
                            P.NroPed ASC,
                            P.FeSolEnt ASC,
                            P.NuOrItem ASC", nroClieParam, nroPedParam)
                    .ToListAsync();

                _logger.LogInformation($"Se encontraron {pedidosPendientes.Count} items pendientes para cliente {numeroCliente}");

                return Ok(pedidosPendientes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener pedidos pendientes para cliente {numeroCliente}");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Método de diagnóstico para verificar datos en la tabla de auditoría
        /// </summary>
        [HttpGet("auditoria/{nroPed}/diagnostico")]
        public async Task<ActionResult<object>> DiagnosticoAuditoria(int nroPed)
        {
            try
            {
                var totalRegistros = await _context.PedidosAuditoria.CountAsync();
                var registrosPedido = await _context.PedidosAuditoria
                    .Where(a => a.NroPed == nroPed)
                    .CountAsync();
                var registrosPeddeta = await _context.PedidosAuditoria
                    .Where(a => a.NroPed == nroPed && a.Tabla == "PEDDETA")
                    .CountAsync();
                var registrosInsert = await _context.PedidosAuditoria
                    .Where(a => a.NroPed == nroPed && a.Tabla == "PEDDETA" && a.Accion == "INSERT")
                    .CountAsync();

                var muestraRegistros = await _context.PedidosAuditoria
                    .Where(a => a.NroPed == nroPed)
                    .Take(3)
                    .Select(a => new {
                        a.IdAuditoria,
                        a.Tabla,
                        a.Accion,
                        a.Campo,
                        TieneValorNuevo = !string.IsNullOrEmpty(a.ValorNuevo),
                        LongitudValorNuevo = a.ValorNuevo != null ? a.ValorNuevo.Length : 0,
                        a.FechaCambio
                    })
                    .ToListAsync();

                return Ok(new
                {
                    TotalRegistros = totalRegistros,
                    RegistrosPedido = registrosPedido,
                    RegistrosPeddeta = registrosPeddeta,
                    RegistrosInsert = registrosInsert,
                    MuestraRegistros = muestraRegistros
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error en diagnóstico para pedido {nroPed}");
                return StatusCode(500, ex.Message);
            }
        }

        /// <summary>
        /// Obtiene la auditoría de PEDDETA para un pedido específico
        /// Parsea el campo ValorNuevo para extraer los valores individuales
        /// </summary>
        /// <param name="nroPed">Número de pedido</param>
        /// <returns>Lista de registros de auditoría parseados</returns>
        [HttpGet("auditoria/{nroPed}")]
        public async Task<ActionResult<List<PeddetaAuditoriaDTO>>> ObtenerAuditoriaPeddeta(int nroPed)
        {
            try
            {
                _logger.LogInformation($"Obteniendo auditoría de PEDDETA para pedido: {nroPed}");

                // Consulta SQL directa para obtener los registros de auditoría
                // Agregamos validaciones para manejar valores NULL
                var registrosAuditoria = await _context.PedidosAuditoria
                    .Where(a => a.NroPed == nroPed && 
                               a.Tabla == "PEDDETA" && 
                               a.Accion == "INSERT" &&
                               a.ValorNuevo != null) // Solo registros con ValorNuevo no nulo
                    .OrderBy(a => a.FechaCambio)
                    .ThenBy(a => a.IdAuditoria)
                    .ToListAsync();

                var resultado = new List<PeddetaAuditoriaDTO>();

                foreach (var registro in registrosAuditoria)
                {
                    try
                    {
                        var dto = new PeddetaAuditoriaDTO
                        {
                            IdAuditoria = registro.IdAuditoria,
                            NroPed = registro.NroPed,
                            Accion = registro.Accion ?? "INSERT", // Valor por defecto si es NULL
                            FechaCambio = registro.FechaCambio,
                            NumUsuario = ParseNumUsuario(registro.Usuario)
                        };

                        // Parsear el campo ValorNuevo para extraer los valores
                        if (!string.IsNullOrEmpty(registro.ValorNuevo))
                        {
                            ParsearValorNuevo(registro.ValorNuevo, dto);
                        }

                        resultado.Add(dto);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"Error al procesar registro de auditoría ID: {registro.IdAuditoria}");
                        // Continuar con el siguiente registro en lugar de fallar completamente
                        continue;
                    }
                }

                _logger.LogInformation($"Se encontraron {resultado.Count} registros de auditoría para el pedido {nroPed}");
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener auditoría de PEDDETA para pedido {nroPed}");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Parsea el campo ValorNuevo que tiene formato: "Campo:Valor, Campo:Valor, ..."
        /// </summary>
        private void ParsearValorNuevo(string valorNuevo, PeddetaAuditoriaDTO dto)
        {
            try
            {
                // Dividir por comas y procesar cada par campo:valor
                var pares = valorNuevo.Split(',');

                foreach (var par in pares)
                {
                    var parTrimmed = par.Trim();
                    var separadorIndex = parTrimmed.IndexOf(':');
                    
                    if (separadorIndex > 0)
                    {
                        var campo = parTrimmed.Substring(0, separadorIndex).Trim();
                        var valor = parTrimmed.Substring(separadorIndex + 1).Trim();

                        // Asignar valores según el campo
                        // Usar CultureInfo.InvariantCulture para parsear correctamente los valores con punto decimal
                        switch (campo)
                        {
                            case "CanPed":
                                if (decimal.TryParse(valor, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal canPed))
                                    dto.CanPed = canPed;
                                break;

                            case "PreUnid":
                                if (decimal.TryParse(valor, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal preUnid))
                                    dto.PreUnid = preUnid;
                                break;

                            case "CantEnt":
                                if (decimal.TryParse(valor, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal cantEnt))
                                    dto.CantEnt = cantEnt;
                                break;

                            case "NroOcom":
                                dto.NroOcom = valor;
                                break;

                            case "Vended":
                                // Parsear como decimal y convertir a entero para manejar valores como "2.00"
                                if (decimal.TryParse(valor, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal vendedDecimal))
                                    dto.Vended = (int)vendedDecimal;
                                break;

                            case "NuOrItem":
                                // Parsear como decimal y convertir a entero para manejar valores como "1.00"
                                if (decimal.TryParse(valor, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal nuOrItemDecimal))
                                    dto.NuOrItem = (int)nuOrItemDecimal;
                                break;

                            case "CodiInt":
                                // Parsear como decimal y convertir a entero para manejar valores como "793.00"
                                if (decimal.TryParse(valor, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal codiIntDecimal))
                                    dto.CodiInt = (int)codiIntDecimal;
                                break;

                            case "CodiExt":
                                dto.CodiExt = valor;
                                break;

                            case "Descrip":
                                dto.Descrip = valor;
                                break;

                            case "MoneEmis":
                                // Parsear como decimal y convertir a entero para manejar valores como "1.00"
                                if (decimal.TryParse(valor, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal moneEmisDecimal))
                                    dto.MoneEmis = (int)moneEmisDecimal;
                                break;

                            case "TipoCam":
                                if (decimal.TryParse(valor, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal tipoCam))
                                    dto.TipoCam = tipoCam;
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"Error al parsear ValorNuevo: {valorNuevo}");
            }
        }

        /// <summary>
        /// Parsea el campo Usuario para extraer el NumUsuario si está disponible
        /// </summary>
        private int? ParseNumUsuario(string? usuario)
        {
            try
            {
                if (string.IsNullOrEmpty(usuario))
                    return null;

                // Intentar extraer número del campo usuario
                var match = Regex.Match(usuario, @"\d+");
                if (match.Success && int.TryParse(match.Value, out int numUsuario))
                {
                    return numUsuario;
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"Error al parsear usuario: {usuario}");
                return null;
            }
        }

        /// <summary>
        /// Obtiene estadísticas de auditoría para un pedido
        /// </summary>
        [HttpGet("auditoria/{nroPed}/estadisticas")]
        public async Task<ActionResult<object>> ObtenerEstadisticasAuditoria(int nroPed)
        {
            try
            {
                var estadisticas = await _context.PedidosAuditoria
                    .Where(a => a.NroPed == nroPed && 
                               a.Tabla == "PEDDETA" && 
                               a.Accion != null) // Filtrar valores NULL
                    .GroupBy(a => a.Accion)
                    .Select(g => new { Accion = g.Key ?? "UNKNOWN", Cantidad = g.Count() })
                    .ToListAsync();

                var totalRegistros = await _context.PedidosAuditoria
                    .CountAsync(a => a.NroPed == nroPed && a.Tabla == "PEDDETA");

                var fechaUltimoCambio = await _context.PedidosAuditoria
                    .Where(a => a.NroPed == nroPed && a.Tabla == "PEDDETA")
                    .MaxAsync(a => (DateTime?)a.FechaCambio);

                return Ok(new
                {
                    NroPedido = nroPed,
                    TotalRegistros = totalRegistros,
                    FechaUltimoCambio = fechaUltimoCambio,
                    PorAccion = estadisticas
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener estadísticas de auditoría para pedido {nroPed}");
                return StatusCode(500, "Error interno del servidor");
            }
        }
        /// <summary>
        /// Obtiene pedidos pendientes filtrados opcionalmente por cliente
        /// Usado en modal "Asignado a Pedidos" de Stock
        /// </summary>
        /// <param name = "nroClie" > Número de cliente(opcional, 0 = todos)</param>
        /// <returns>Lista de pedidos pendientes</returns>
        [HttpGet("pedidos-pendientes-por-cliente")]
        public async Task<ActionResult<List<PedidoPendienteClienteDTO>>> GetPedidosPendientesPorCliente([FromQuery] int nroClie = 0)
        {
            try
            {
                var query = _context.Pedencas
                    .AsNoTracking()
                    .Where(p => p.Status < 8 && p.Status >= 0)
                    .Where(p => _context.Peddetas
                        .Any(pp =>
                            pp.NroPed == p.NroPed &&
                            pp.Status < 8 && pp.Status >= 0 &&
                            ((pp.CanPed - pp.CantEnt) >= 0.05m ||
                             (pp.TipoPed == "ABIERTO" && pp.FinVigen >= DateTime.Now)) &&
                            (nroClie == 0 || pp.NroClie == nroClie)
                        )
                    )
                    .Select(p => new PedidoPendienteClienteDTO
                    {
                        NroPed = p.NroPed ?? 0,
                        Referencia = p.Referencia ?? ""
                    })
                    .OrderByDescending(x => x.NroPed);

                var result = await query.ToListAsync();
                return Ok(result);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener pedidos pendientes: {ex.Message}");
                return StatusCode(500, $"Error al obtener pedidos pendientes: {ex.Message}");
            }
        }
        /// <summary>
        /// Obtiene el detalle de productos pendientes de un pedido específico
        /// Usado en modal "Asignado a Pedidos" de Stock
        /// </summary>
        /// <param name="nroPed">Número de pedido</param>
        /// <returns>Lista de productos pendientes del pedido</returns>
        [HttpGet("detalle-pendiente/{nroPed:int}")]
        public async Task<ActionResult<List<DetallePedidoPendienteDTO>>> GetDetallePendiente(int nroPed)
        {
            try
            {
                if (nroPed < 1)
                {
                    return BadRequest("Número de pedido inválido");
                }

                var hoy = DateTime.Now.Date;

                var detalle = await (from d in _context.Peddetas.AsNoTracking()
                                     join m in _context.masters.AsNoTracking()
                                         on d.CodiInt equals m.Codint into masterJoin
                                     from master in masterJoin.DefaultIfEmpty()
                                     where d.NroPed == nroPed
                                           && d.Status >= 0 && d.Status < 8
                                           && ((d.CanPed ?? 0) - (d.CantEnt ?? 0)) > 0
                                     orderby d.NuOrItem
                                     select new
                                     {
                                         CodiInt = d.CodiInt ?? 0,
                                         CodiExt = d.CodiExt ?? "",
                                         Descrip = d.Descrip ?? "",
                                         CanPed = d.CanPed ?? 0,
                                         CantEnt = d.CantEnt ?? 0,
                                         NroPed = d.NroPed ?? 0,
                                         FeSolEnt = d.FeSolEnt,
                                         NuOrItem = d.NuOrItem ?? 0,
                                         PeddetaId = d.PeddetaId,
                                         Atraso = d.FeSolEnt.HasValue && d.FeSolEnt.Value < hoy
                                             ? (int)(hoy - d.FeSolEnt.Value).TotalDays
                                             : 0,
                                         PendienteRemitir = (d.CanPed ?? 0) - (d.CantEnt ?? 0),
                                         Trazable = master != null ? (master.Trazable ?? 0) : (short)0
                                     })
                    .ToListAsync();

                // Mapear a DTO e inicializar CantidadARemitir
                var resultado = detalle.Select(d => new DetallePedidoPendienteDTO
                {
                    CodiInt = d.CodiInt,
                    CodiExt = d.CodiExt,
                    Descrip = d.Descrip,
                    CanPed = d.CanPed,
                    CantEnt = d.CantEnt,
                    NroPed = d.NroPed,
                    FeSolEnt = d.FeSolEnt,
                    NuOrItem = d.NuOrItem,
                    PeddetaId = d.PeddetaId,
                    Atraso = d.Atraso,
                    PendienteRemitir = d.PendienteRemitir,
                    CantidadARemitir = d.PendienteRemitir, // Inicializar con el pendiente
                    Trazable = d.Trazable,
                    LoteSeleccionado = null
                }).ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener detalle pendiente del pedido {nroPed}: {ex.Message}");
                return StatusCode(500, $"Error al obtener detalle pendiente: {ex.Message}");
            }
        }

    }
}
