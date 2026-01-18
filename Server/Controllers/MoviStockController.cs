using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para manejar movimientos de stock
    /// Expone la funcionalidad del servicio MoviStockService a través de API REST
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme)]
    public class MoviStockController : ControllerBase
    {
        private readonly IMoviStockService _moviStockService;
        private readonly ILogger<MoviStockController> _logger;

        public MoviStockController(
            IMoviStockService moviStockService,
            ILogger<MoviStockController> logger)
        {
            _moviStockService = moviStockService;
            _logger = logger;
        }

        /// <summary>
        /// Registra un movimiento de stock
        /// Equivalente a la llamada MOVISTOK en VB6
        /// </summary>
        /// <param name="movimiento">Datos del movimiento de stock</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPost("registrar")]
        public async Task<ActionResult> RegistrarMovimiento([FromBody] MoviStockDTO movimiento)
        {
            try
            {
                _logger.LogInformation("Iniciando registro de movimiento de stock para producto {CodigoInterno}", 
                    movimiento.CodigoInterno);

                var resultado = await _moviStockService.RegistrarMovimientoAsync(movimiento, User);
                
                if (resultado is OkObjectResult)
                {
                    _logger.LogInformation("Movimiento de stock registrado exitosamente para producto {CodigoInterno}", 
                        movimiento.CodigoInterno);
                }
                else
                {
                    _logger.LogWarning("Error al registrar movimiento de stock para producto {CodigoInterno}", 
                        movimiento.CodigoInterno);
                }

                return resultado;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inesperado al registrar movimiento de stock");
                return BadRequest($"Error inesperado: {ex.Message}");
            }
        }

        /// <summary>
        /// Valida un movimiento de stock sin registrarlo
        /// Útil para validaciones previas en el frontend
        /// </summary>
        /// <param name="movimiento">Datos del movimiento a validar</param>
        /// <returns>True si el movimiento es válido, False en caso contrario</returns>
        [HttpPost("validar")]
        public async Task<ActionResult<bool>> ValidarMovimiento([FromBody] MoviStockDTO movimiento)
        {
            try
            {
                _logger.LogInformation("Validando movimiento de stock para producto {CodigoInterno}", 
                    movimiento.CodigoInterno);

                bool esValido = await _moviStockService.ValidarMovimientoAsync(movimiento);
                
                return Ok(esValido);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al validar movimiento de stock");
                return BadRequest($"Error al validar movimiento: {ex.Message}");
            }
        }

        /// <summary>
        /// Valida si existe un remito duplicado (mismo prefijo y número)
        /// </summary>
        /// <param name="prefijo">Prefijo del remito (NUMPVTA)</param>
        /// <param name="numero">Número del remito (DOPRINUM)</param>
        /// <returns>True si existe, False si no existe</returns>
        [HttpGet("validar-remito/{prefijo}/{numero}")]
        public async Task<ActionResult<bool>> ValidarRemitoDuplicado(short prefijo, int numero)
        {
            try
            {
                _logger.LogInformation("Validando si existe remito {Prefijo}-{Numero}", prefijo, numero);

                bool existe = await _moviStockService.ExisteRemitoDuplicadoAsync(prefijo, numero);
                
                return Ok(existe);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al validar remito duplicado");
                return BadRequest($"Error al validar remito: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene los remitos de un mes específico para anulación
        /// </summary>
        /// <param name="anio">Año</param>
        /// <param name="mes">Mes (1-12)</param>
        /// <param name="numeroCliente">Número de cliente (opcional)</param>
        /// <returns>Lista de remitos del mes</returns>
        [HttpGet("remitos-mes/{anio}/{mes}")]
        public async Task<ActionResult<List<RemitoAnularDTO>>> ObtenerRemitosMes(int anio, int mes, [FromQuery] int? numeroCliente = null)
        {
            try
            {
                _logger.LogInformation("Obteniendo remitos del mes {Mes}/{Anio}, Cliente: {Cliente}", 
                    mes, anio, numeroCliente?.ToString() ?? "Todos");

                var remitos = await _moviStockService.ObtenerRemitosMesAsync(anio, mes, numeroCliente);
                
                return Ok(remitos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener remitos del mes");
                return BadRequest($"Error al obtener remitos: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene todos los remitos de un mes específico para consulta (incluye anulados)
        /// </summary>
        /// <param name="anio">Año</param>
        /// <param name="mes">Mes (1-12)</param>
        /// <param name="numeroCliente">Número de cliente (opcional)</param>
        /// <returns>Lista de remitos del mes, incluyendo los anulados</returns>
        [HttpGet("remitos-mes-todos/{anio}/{mes}")]
        public async Task<ActionResult<List<RemitoAnularDTO>>> ObtenerRemitosMesTodos(int anio, int mes, [FromQuery] int? numeroCliente = null)
        {
            try
            {
                _logger.LogInformation("Obteniendo TODOS los remitos del mes {Mes}/{Anio}, Cliente: {Cliente}",
                    mes, anio, numeroCliente?.ToString() ?? "Todos");

                var remitos = await _moviStockService.ObtenerRemitosMesTodosAsync(anio, mes, numeroCliente);

                return Ok(remitos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener TODOS los remitos del mes");
                return BadRequest($"Error al obtener remitos: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el último número de remito registrado en MOVISTO
        /// Filtra por código de movimiento RT y fecha mayor o igual a la fecha de nueva numeración
        /// </summary>
        /// <param name="fechaNuevaNumeracion">Fecha desde la cual considerar los remitos (opcional, formato ISO: yyyy-MM-dd)</param>
        /// <returns>Último número de remito</returns>
        [HttpGet("ultimo-numero-remito")]
        public async Task<ActionResult<int>> ObtenerUltimoNumeroRemito([FromQuery] string? fechaNuevaNumeracion = null)
        {
            try
            {
                DateTime? fecha = null;
                if (!string.IsNullOrEmpty(fechaNuevaNumeracion))
                {
                    if (DateTime.TryParse(fechaNuevaNumeracion, out DateTime parsedFecha))
                    {
                        fecha = parsedFecha;
                        _logger.LogInformation("Obteniendo último número de remito desde fecha: {Fecha:dd/MM/yyyy}", fecha);
                    }
                }
                else
                {
                    _logger.LogInformation("Obteniendo último número de remito (sin filtro de fecha)");
                }

                var ultimoNumero = await _moviStockService.ObtenerUltimoNumeroRemitoAsync(fecha);
                
                _logger.LogInformation("Último número de remito: {UltimoNumero}", ultimoNumero);
                return Ok(ultimoNumero);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener último número de remito");
                return BadRequest($"Error al obtener último número de remito: {ex.Message}");
            }
        }

        /// <summary>
        /// Anula un remito específico
        /// Replica la funcionalidad de anulación de remitos de VB6
        /// </summary>
        /// <param name="request">Datos del remito a anular</param>
        /// <returns>Resultado de la anulación</returns>
        [HttpPost("anular-remito")]
        public async Task<ActionResult<AnularRemitoResponse>> AnularRemito([FromBody] AnularRemitoRequest request)
        {
            try
            {
                _logger.LogInformation("Solicitud de anulación de remito {NumeroRemito} para cliente {NumeroCliente}", 
                    request.NumeroRemito, request.NumeroCliente);

                var resultado = await _moviStockService.AnularRemitoAsync(request);
                
                if (resultado.Exitoso)
                {
                    return Ok(resultado);
                }
                else
                {
                    return BadRequest(resultado);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al anular remito");
                return BadRequest(new AnularRemitoResponse
                {
                    Exitoso = false,
                    Mensaje = $"Error al anular remito: {ex.Message}"
                });
            }
        }
    }
}
