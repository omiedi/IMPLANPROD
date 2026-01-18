using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para manejar operaciones con la tabla ORFAPEN
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class OrfapenController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<OrfapenController> _logger;

        public OrfapenController(DataContext context, ILogger<OrfapenController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene el depósito de entrada para una orden de fabricación específica
        /// </summary>
        /// <param name="ordenFabricacion">Número de orden de fabricación</param>
        /// <param name="codigoMovimiento">Código de movimiento (ej: 'OF')</param>
        /// <returns>Código del depósito de entrada o null si no se encuentra</returns>
        [HttpGet("deposito-entrada")]
        public async Task<ActionResult<int?>> GetDepositoEntrada(
            [FromQuery] string ordenFabricacion, 
            [FromQuery] string codigoMovimiento)
        {
            try
            {
                // Validar parámetros de entrada
                if (string.IsNullOrEmpty(ordenFabricacion) || string.IsNullOrEmpty(codigoMovimiento))
                {
                    _logger.LogWarning("Parámetros vacíos o nulos: ordenFabricacion={OrdenFabricacion}, codigoMovimiento={CodigoMovimiento}", 
                        ordenFabricacion, codigoMovimiento);
                    return BadRequest("Los parámetros ordenFabricacion y codigoMovimiento son requeridos");
                }

                _logger.LogInformation("Buscando depósito de entrada para O.F. {OrdenFabricacion} con código {CodigoMovimiento}", 
                    ordenFabricacion, codigoMovimiento);

                // Buscar en ORFAPEN donde DOPRICOR = ordenFabricacion AND CODIMOVI = codigoMovimiento
                var deposito = await _context.Orfapens
                    .Where(o => o.DoPriCor == ordenFabricacion && o.CodiMovi == codigoMovimiento)
                    .Select(o => o.DepoMovi)
                    .FirstOrDefaultAsync();

                if (!deposito.HasValue)
                {
                    _logger.LogWarning("No se encontró depósito para O.F. {OrdenFabricacion} con código {CodigoMovimiento}", 
                        ordenFabricacion, codigoMovimiento);
                    return NotFound($"No se encontró depósito de entrada para la orden {ordenFabricacion} con código {codigoMovimiento}");
                }

                _logger.LogInformation("Depósito de entrada encontrado: {Deposito} para O.F. {OrdenFabricacion}", 
                    deposito.Value, ordenFabricacion);

                return Ok(deposito.Value);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener depósito de entrada para O.F. {OrdenFabricacion}", ordenFabricacion);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene los datos de ORFAPEN necesarios para la aprobación
        /// Replica la consulta SQL del código VB6
        /// </summary>
        /// <param name="ordenFabricacion">Número de orden de fabricación</param>
        /// <param name="idTexto">ID de texto del cierre</param>
        /// <returns>Datos completos de ORFAPEN para aprobación</returns>
        [HttpGet("datos-aprobacion")]
        public async Task<ActionResult<DatosOrfapenDTO>> GetDatosAprobacion(
            [FromQuery] string ordenFabricacion, 
            [FromQuery] string idTexto)
        {
            try
            {
                if (string.IsNullOrEmpty(ordenFabricacion) || string.IsNullOrEmpty(idTexto))
                {
                    return BadRequest("Los parámetros ordenFabricacion e idTexto son requeridos");
                }

                _logger.LogInformation("Obteniendo datos de aprobación para O.F. {OrdenFabricacion} con ID {IdTexto}", 
                    ordenFabricacion, idTexto);

                // Consulta simplificada - primero obtener datos de ORFAPEN
                var orfapen = await _context.Orfapens
                    .Where(o => o.DoPriCor == ordenFabricacion 
                             && o.CodiMovi == "OF" 
                             && o.IDTEXT == idTexto 
                             && o.CantIngre > 0)
                    .OrderBy(o => o.FeReMovi)
                    .FirstOrDefaultAsync();

                if (orfapen == null)
                {
                    _logger.LogWarning("No se encontraron datos de ORFAPEN para O.F. {OrdenFabricacion}", ordenFabricacion);
                    return NotFound($"No se encontraron datos para la orden {ordenFabricacion}");
                }

                // Obtener datos del master por separado
                var master = await _context.masters
                    .Where(m => m.Codint == orfapen.Cod_Inte)
                    .FirstOrDefaultAsync();

                var datos = new DatosOrfapenDTO
                {
                    ReferCor = orfapen.ReferCor ?? string.Empty,
                    FechMov = orfapen.FechMov ?? DateTime.Now,
                    CodInte = orfapen.Cod_Inte ?? 0,
                    CodExte = orfapen.Cod_Exte ?? string.Empty,
                    CodiMovi = orfapen.CodiMovi ?? string.Empty,
                    DoPriCor = orfapen.DoPriCor ?? string.Empty,
                    DoSecCor = orfapen.DoSecCor ?? string.Empty,
                    DepoMovi = orfapen.DepoMovi ?? 0,
                    DepoConsu = orfapen.DepoConsu ?? 0,
                    Cantidad = orfapen.CantIngre ?? 0,
                    CodigoExterno = master?.Codigo ?? string.Empty,
                    Descripcion = master?.Descripcion ?? string.Empty
                };

                if (datos == null)
                {
                    _logger.LogWarning("No se encontraron datos de aprobación para O.F. {OrdenFabricacion}", ordenFabricacion);
                    return NotFound($"No se encontraron datos para la orden {ordenFabricacion}");
                }

                // Formatear DoPriCor y DoSecCor como en VB6 (rellenar con ceros)
                datos.DoPriCor = FormatearDocumento(datos.DoPriCor);
                datos.DoSecCor = FormatearDocumento(datos.DoSecCor);

                _logger.LogInformation("Datos de aprobación obtenidos para O.F. {OrdenFabricacion}", ordenFabricacion);
                return Ok(datos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener datos de aprobación para O.F. {OrdenFabricacion}", ordenFabricacion);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Formatea un documento agregando ceros como en VB6
        /// While Len(DOPRICO$) < 11: DOPRICO$ = Mid$(DOPRICO$, 1, 3) & "0" & Mid$(DOPRICO$, 4): Wend
        /// </summary>
        private string FormatearDocumento(string documento)
        {
            if (string.IsNullOrEmpty(documento) || documento.Length >= 11)
                return documento;

            while (documento.Length < 11)
            {
                if (documento.Length >= 3)
                {
                    documento = documento.Substring(0, 3) + "0" + documento.Substring(3);
                }
                else
                {
                    documento = "0" + documento;
                }
            }

            return documento;
        }

        /// <summary>
        /// Obtiene todos los registros de ORFAPEN para una orden de fabricación específica
        /// </summary>
        /// <param name="ordenFabricacion">Número de orden de fabricación</param>
        /// <returns>Lista de registros ORFAPEN</returns>
        [HttpGet("por-orden/{ordenFabricacion}")]
        public async Task<ActionResult<List<Orfapen>>> GetPorOrdenFabricacion(string ordenFabricacion)
        {
            try
            {
                if (string.IsNullOrEmpty(ordenFabricacion))
                {
                    return BadRequest("El número de orden de fabricación es requerido");
                }

                var registros = await _context.Orfapens
                    .Where(o => o.DoPriCor == ordenFabricacion)
                    .OrderBy(o => o.FechMov)
                    .ToListAsync();

                return Ok(registros);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener registros ORFAPEN para O.F. {ordenFabricacion}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene registros de ORFAPEN filtrados por múltiples criterios
        /// </summary>
        /// <param name="codigoMovimiento">Código de movimiento</param>
        /// <param name="fechaDesde">Fecha desde (opcional)</param>
        /// <param name="fechaHasta">Fecha hasta (opcional)</param>
        /// <returns>Lista de registros ORFAPEN filtrados</returns>
        [HttpGet("filtrados")]
        public async Task<ActionResult<List<Orfapen>>> GetFiltrados(
            [FromQuery] string? codigoMovimiento = null,
            [FromQuery] DateTime? fechaDesde = null,
            [FromQuery] DateTime? fechaHasta = null)
        {
            try
            {
                var query = _context.Orfapens.AsQueryable();

                // Aplicar filtros
                if (!string.IsNullOrEmpty(codigoMovimiento))
                {
                    query = query.Where(o => o.CodiMovi == codigoMovimiento);
                }

                if (fechaDesde.HasValue)
                {
                    query = query.Where(o => o.FechMov >= fechaDesde.Value);
                }

                if (fechaHasta.HasValue)
                {
                    query = query.Where(o => o.FechMov <= fechaHasta.Value);
                }

                var registros = await query
                    .OrderByDescending(o => o.FechMov)
                    .Take(1000) // Limitar resultados para evitar sobrecarga
                    .ToListAsync();

                return Ok(registros);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener registros ORFAPEN filtrados");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el depósito de consumo (DEPOCONSU) desde ORFAPEN
        /// donde DoPriCor = ordenFabricacion AND CODIMOVI = 'CF'
        /// y valida que el depósito exista y esté activo en TADEPOSI
        /// </summary>
        /// <param name="ordenFabricacion">Número de orden de fabricación</param>
        /// <returns>Información del depósito de consumo con estado de activación</returns>
        [HttpGet("deposito-consumo")]
        public async Task<ActionResult<DepositoConsumoResponseDTO>> GetDepositoConsumo([FromQuery] string ordenFabricacion)
        {
            try
            {
                // Validar parámetro de entrada
                if (string.IsNullOrEmpty(ordenFabricacion))
                {
                    _logger.LogWarning("Parámetro ordenFabricacion vacío o nulo");
                    return BadRequest("El número de orden de fabricación es requerido");
                }

                _logger.LogInformation($"Buscando depósito de consumo para O.F. {ordenFabricacion} con código CF");

                // Buscar en ORFAPEN donde DOPRICOR = ordenFabricacion AND CODIMOVI = 'CF'
                var registro = await _context.Orfapens
                    .Where(o => o.DoPriCor == ordenFabricacion && o.CodiMovi == "CF")
                    .Select(o => o.DepoConsu)
                    .FirstOrDefaultAsync();

                if (!registro.HasValue)
                {
                    _logger.LogWarning($"No se encontró registro en ORFAPEN para O.F. {ordenFabricacion} con código CF");
                    return NotFound($"No se encontró depósito de consumo para la orden {ordenFabricacion}");
                }

                var depositoConsumo = registro.Value;
                _logger.LogInformation($"Depósito de consumo encontrado: {depositoConsumo} para O.F. {ordenFabricacion}");

                // Buscar información del depósito en TADEPOSI
                var depositoInfo = await _context.Tadeposi
                    .Where(d => d.CodigoDepo == depositoConsumo)
                    .FirstOrDefaultAsync();

                var response = new DepositoConsumoResponseDTO
                {
                    CodigoDeposito = depositoConsumo
                };

                if (depositoInfo == null)
                {
                    // El depósito no existe en TADEPOSI
                    response.EstaActivo = false;
                    response.Descripcion = "Depósito no encontrado";
                    response.Mensaje = $"El depósito {depositoConsumo} no existe en el sistema";
                    response.RequiereSeleccion = true;
                    _logger.LogWarning($"Depósito {depositoConsumo} no existe en TADEPOSI");
                }
                else if (depositoInfo.Status < 0)
                {
                    // El depósito existe pero está inactivo
                    response.EstaActivo = false;
                    response.Descripcion = depositoInfo.Descripcion ?? "Sin descripción";
                    response.Mensaje = $"El depósito {depositoConsumo} está inactivo";
                    response.RequiereSeleccion = true;
                    _logger.LogWarning($"Depósito {depositoConsumo} está inactivo en TADEPOSI");
                }
                else
                {
                    // El depósito existe y está activo
                    response.EstaActivo = true;
                    response.Descripcion = depositoInfo.Descripcion ?? "Sin descripción";
                    response.Mensaje = $"Depósito {depositoConsumo} activo y válido";
                    response.RequiereSeleccion = false;
                    _logger.LogInformation($"Depósito {depositoConsumo} validado como activo en TADEPOSI");
                }

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener depósito de consumo para O.F. {ordenFabricacion}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }
    }
}
