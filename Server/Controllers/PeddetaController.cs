using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
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

        /// <summary>
        /// Obtiene los items de pedidos pendientes de entrega para un cliente
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
                        AND CanSaldo IS NOT NULL 
                        AND CanSaldo > 0 
                        AND Status = 0
                        ORDER BY NroPed, FeSolEnt", numeroCliente)
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
    }
}
