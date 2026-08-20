using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controller para gestionar costos de items de órdenes de compra
    /// Requiere derecho "VerCostosOC" para acceder
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class OCOMDETACostoController : ControllerBase
    {
        private readonly IOCOMDETACostoService _service;

        public OCOMDETACostoController(IOCOMDETACostoService service)
        {
            _service = service;
        }

        /// <summary>
        /// Obtiene los costos de items de una orden de compra específica
        /// GET: api/OCOMDETACosto/orden/12345
        /// </summary>
        /// <param name="numeOcom">Número de orden de compra</param>
        /// <returns>Lista de items con sus costos</returns>
        [HttpGet("orden/{numeOcom}")]
        public async Task<ActionResult<List<OCOMDETACostoDto>>> GetCostosPorOrden(int numeOcom)
        {
            try
            {
                var costos = await _service.GetCostosPorOrdenAsync(numeOcom);
                return Ok(costos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error al obtener costos: {ex.Message}" });
            }
        }

        /// <summary>
        /// Guarda o actualiza el costo de un item
        /// POST: api/OCOMDETACosto
        /// </summary>
        /// <param name="dto">Datos del costo a guardar</param>
        /// <returns>DTO actualizado con el ID</returns>
        [HttpPost]
        public async Task<ActionResult<OCOMDETACostoDto>> GuardarCosto([FromBody] OCOMDETACostoDto dto)
        {
            try
            {
                if (dto == null)
                {
                    return BadRequest(new { message = "Los datos del costo son requeridos" });
                }

                if (!string.IsNullOrEmpty(dto.Observaciones) && dto.Observaciones.Length > 1024)
                {
                    return BadRequest(new { message = "El campo Observaciones no puede superar los 1024 caracteres." });
                }

                var costoGuardado = await _service.GuardarCostoAsync(dto);
                return Ok(costoGuardado);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error al guardar costo: {ex.Message}" });
            }
        }

        /// <summary>
        /// Guarda o actualiza múltiples costos en una sola operación
        /// POST: api/OCOMDETACosto/batch
        /// </summary>
        /// <param name="dtos">Lista de costos a guardar</param>
        /// <returns>Lista de DTOs actualizados</returns>
        [HttpPost("batch")]
        public async Task<ActionResult<List<OCOMDETACostoDto>>> GuardarCostos([FromBody] List<OCOMDETACostoDto> dtos)
        {
            try
            {
                if (dtos == null || !dtos.Any())
                {
                    return BadRequest(new { message = "La lista de costos no puede estar vacía" });
                }

                if (dtos.Any(x => !string.IsNullOrEmpty(x.Observaciones) && x.Observaciones.Length > 1024))
                {
                    return BadRequest(new { message = "Una o más observaciones superan los 1024 caracteres." });
                }

                var costosGuardados = await _service.GuardarCostosAsync(dtos);
                return Ok(costosGuardados);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error al guardar costos: {ex.Message}" });
            }
        }

        /// <summary>
        /// Guarda costos y registra cambios en NOVEDADES
        /// POST: api/OCOMDETACosto/batch-con-novedades
        /// Si la configuración ACOSREP es diferente de "NUNCA", registra los cambios de precio en NOVEDADES
        /// </summary>
        /// <param name="dto">Datos de costos, observación y configuración ACOSREP</param>
        /// <returns>Lista de DTOs actualizados</returns>
        [HttpPost("batch-con-novedades")]
        public async Task<ActionResult<List<OCOMDETACostoDto>>> GuardarCostosConNovedades([FromBody] GuardarCostosConNovedadesDto dto)
        {
            try
            {
                if (dto == null || dto.Costos == null || !dto.Costos.Any())
                {
                    return BadRequest(new { message = "La lista de costos no puede estar vacía" });
                }

                if (dto.Costos.Any(x => !string.IsNullOrEmpty(x.Observaciones) && x.Observaciones.Length > 1024))
                {
                    return BadRequest(new { message = "Una o más observaciones superan los 1024 caracteres." });
                }

                var costosGuardados = await _service.GuardarCostosConNovedadesAsync(dto);
                return Ok(costosGuardados);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error al guardar costos con novedades: {ex.Message}" });
            }
        }
    }
}
