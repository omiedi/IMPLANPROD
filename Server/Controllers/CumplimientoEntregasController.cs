using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Control de Cumplimiento de Entregas (migración de COCUENT09.FRM - COCUME09).
    /// IC-R: índice por remito (aplica tolerancia). IC-P: índice por pedido.
    /// </summary>
    [ApiController]
    [Route("api/cumplimiento-entregas")]
    public class CumplimientoEntregasController : ControllerBase
    {
        private readonly CumplimientoEntregasService _service;
        private readonly ILogger<CumplimientoEntregasController> _logger;

        public CumplimientoEntregasController(CumplimientoEntregasService service, ILogger<CumplimientoEntregasController> logger)
        {
            _service = service;
            _logger = logger;
        }

        /// <summary>
        /// GET api/cumplimiento-entregas?tipo=REMITO|PEDIDO&amp;fechaDesde=yyyy-MM-dd&amp;fechaHasta=yyyy-MM-dd&amp;clienteDesde=&amp;clienteHasta=
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<CumplimientoEntregaResultadoDTO>> Get(
            [FromQuery] string tipo,
            [FromQuery] DateTime fechaDesde,
            [FromQuery] DateTime fechaHasta,
            [FromQuery] int? clienteDesde = null,
            [FromQuery] int? clienteHasta = null)
        {
            if (fechaHasta.Date < fechaDesde.Date)
                return BadRequest("La fecha hasta no puede ser anterior a la fecha desde.");

            try
            {
                return Ok(await _service.ObtenerInformeAsync(tipo, fechaDesde, fechaHasta, clienteDesde, clienteHasta));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener cumplimiento de entregas");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        [HttpGet("tolerancia")]
        public async Task<ActionResult<ToleranciaEntregaDTO>> GetTolerancia()
        {
            try
            {
                return Ok(new ToleranciaEntregaDTO { Dias = await _service.ObtenerToleranciaAsync() });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener tolerancia de entrega");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        [HttpPut("tolerancia")]
        public async Task<ActionResult> PutTolerancia([FromBody] ToleranciaEntregaDTO dto)
        {
            if (dto == null || dto.Dias < 0 || dto.Dias > CumplimientoEntregasService.ToleranciaMaxima)
                return BadRequest("Valor Ingresado No Válido (Valores Entre 0 y 30)");

            try
            {
                var ok = await _service.GrabarToleranciaAsync(dto.Dias);
                return ok ? NoContent() : StatusCode(500, "No se pudo grabar la tolerancia");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al grabar tolerancia de entrega");
                return StatusCode(500, "Error interno del servidor");
            }
        }
    }
}
