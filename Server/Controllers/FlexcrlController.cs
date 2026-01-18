using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para manejar las operaciones de controles de procesos de negocio (FLEXCRL)
    /// </summary>
    [ApiController]
    [Route("/api/flexcrl")]
    public class FlexcrlController : ControllerBase
    {
        private readonly FlexcrlService _flexcrlService;
        private readonly IUsuarioFlexoftService _usuarioFlexoftService;
        private readonly ICodiempresa _codiempresa;

        /// <summary>
        /// Constructor del controlador de FLEXCRL
        /// </summary>
        /// <param name="flexcrlService">Servicio de FLEXCRL</param>
        /// <param name="usuarioFlexoftService">Servicio de usuario</param>
        /// <param name="codiempresa">Servicio de código de empresa</param>
        public FlexcrlController(FlexcrlService flexcrlService, IUsuarioFlexoftService usuarioFlexoftService, ICodiempresa codiempresa)
        {
            _flexcrlService = flexcrlService;
            _usuarioFlexoftService = usuarioFlexoftService;
            _codiempresa = codiempresa;
        }

        /// <summary>
        /// Obtiene un control específico por CLABUS y PROBUS
        /// </summary>
        /// <param name="clabus">Clasificación del proceso de negocio</param>
        /// <param name="probus">Proceso de negocio</param>
        /// <returns>Registro FLEXCRL si existe, NotFound si no existe</returns>
        [HttpGet("{clabus}/{probus}")]
        public async Task<ActionResult> ObtenerControl(string clabus, string probus)
        {
            try
            {
                int? idFlexoft = _usuarioFlexoftService.InterObtenerIdFlexoft(User);
                if (!idFlexoft.HasValue)
                {
                    return BadRequest("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Convertir "empty" de vuelta a string vacío para mantener compatibilidad con la lógica VB6
                if (clabus == "empty")
                {
                    clabus = "";
                }

                var control = await _flexcrlService.ObtenerControlAsync(clabus, probus);
                
                if (control == null)
                {
                    return NotFound($"No se encontró control para CLABUS: {clabus}, PROBUS: {probus}");
                }

                return Ok(control);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener el control: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza o crea un control de proceso de negocio
        /// </summary>
        /// <param name="request">Datos del control a actualizar/crear</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPut]
        public async Task<ActionResult> ActualizarOCrearControl([FromBody] FlexcrlRequestDTO request)
        {
            try
            {
                int? idFlexoft = _usuarioFlexoftService.InterObtenerIdFlexoft(User);
                if (!idFlexoft.HasValue)
                {
                    return BadRequest("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                if (request == null)
                {
                    return BadRequest("Los datos del control son requeridos.");
                }

                if (string.IsNullOrWhiteSpace(request.Probus))
                {
                    return BadRequest("El parámetro PROBUS es requerido.");
                }

                // Convertir "empty" de vuelta a string vacío para mantener compatibilidad con la lógica VB6
                string clabus = request.Clabus ?? "";
                if (clabus == "empty")
                {
                    clabus = "";
                }

                string probus = request.Probus;
                if (probus == "empty")
                {
                    probus = "";
                }

                var resultado = await _flexcrlService.ActualizarOCrearControlAsync(
                    clabus, 
                    probus, 
                    request.Valcontrol ?? ""
                );

                if (resultado)
                {
                    return Ok(new { mensaje = "Control actualizado/creado exitosamente." });
                }
                else
                {
                    return BadRequest("No se pudo actualizar/crear el control.");
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al procesar el control: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// DTO para las peticiones de actualización/creación de controles FLEXCRL
    /// </summary>
    public class FlexcrlRequestDTO
    {
        /// <summary>
        /// Clasificación del proceso de negocio
        /// </summary>
        public string? Clabus { get; set; }

        /// <summary>
        /// Proceso de negocio
        /// </summary>
        public string? Probus { get; set; }

        /// <summary>
        /// Valor de control del proceso
        /// </summary>
        public string? Valcontrol { get; set; }
    }
}
