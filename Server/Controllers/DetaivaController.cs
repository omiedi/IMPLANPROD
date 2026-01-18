using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Shared.DTOs;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DetaivaController : ControllerBase
    {
        private readonly DetaivaService _detaivaService;
        private readonly DataContext _context;

        private readonly ILogger<DetaivaController> _logger;

        public DetaivaController(DetaivaService detaivaService, DataContext context, ILogger<DetaivaController> logger)
        {
            _detaivaService = detaivaService;
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene los registros de distribución de IVA para una orden específica por ID de Ocomenca
        /// </summary>
        [HttpGet("orden/{ocomencaId}")]
        public async Task<ActionResult<List<Detaiva>>> GetByOrden(int ocomencaId)
        {
            try
            {
                _logger.LogInformation($"[GetByOrden] Buscando orden con ID: {ocomencaId}");

                // Primero obtener el NumeOcom desde la entidad Ocomenca
                var ocomenca = await _context.Ocomencas
                    .Where(o => o.OCOMENCA_ID == ocomencaId)
                    .Select(o => new { o.NumeOcom })
                    .FirstOrDefaultAsync();

                if (ocomenca == null)
                {
                    _logger.LogWarning($"[GetByOrden] No se encontró orden con ID: {ocomencaId}");
                    return NotFound($"No se encontró la orden de compra con ID {ocomencaId}");
                }

                if (!ocomenca.NumeOcom.HasValue)
                {
                    _logger.LogWarning($"[GetByOrden] Orden ID {ocomencaId} no tiene NumeOcom válido");
                    return BadRequest($"La orden de compra con ID {ocomencaId} no tiene un número de orden válido");
                }

                _logger.LogInformation($"[GetByOrden] Orden ID {ocomencaId} tiene NumeOcom: {ocomenca.NumeOcom.Value}");

                // Ahora buscar los registros DETAIVA usando el NumeOcom
                var registros = await _detaivaService.ObtenerDetalleIvaAsync(ocomenca.NumeOcom.Value);
                
                _logger.LogInformation($"[GetByOrden] Encontrados {registros.Count} registros DETAIVA para NumeOcom: {ocomenca.NumeOcom.Value}");
                
                return Ok(registros);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[GetByOrden] Error al obtener registros DETAIVA para orden ID: {ocomencaId}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene la distribución de IVA como DTO para una orden específica por ID de Ocomenca
        /// Equivalente a cargar RIVACOM07 del código VB6
        /// </summary>
        [HttpGet("distribucion/{ocomencaId}")]
        public async Task<ActionResult<DistribucionIvaDTO>> GetDistribucionByOrden(int ocomencaId)
        {
            try
            {
                _logger.LogInformation($"[GetDistribucionByOrden] Buscando distribución IVA para orden ID: {ocomencaId}");

                // Primero obtener el NumeOcom desde la entidad Ocomenca
                var ocomenca = await _context.Ocomencas
                    .Where(o => o.OCOMENCA_ID == ocomencaId)
                    .Select(o => new { o.NumeOcom })
                    .FirstOrDefaultAsync();

                if (ocomenca == null)
                {
                    _logger.LogWarning($"[GetDistribucionByOrden] No se encontró orden con ID: {ocomencaId}");
                    return NotFound($"No se encontró la orden de compra con ID {ocomencaId}");
                }

                if (!ocomenca.NumeOcom.HasValue)
                {
                    _logger.LogWarning($"[GetDistribucionByOrden] Orden ID {ocomencaId} no tiene NumeOcom válido");
                    return BadRequest($"La orden de compra con ID {ocomencaId} no tiene un número de orden válido");
                }

                _logger.LogInformation($"[GetDistribucionByOrden] Orden ID {ocomencaId} tiene NumeOcom: {ocomenca.NumeOcom.Value}");

                // Ahora buscar los registros DETAIVA usando el NumeOcom
                var registros = await _detaivaService.ObtenerDetalleIvaAsync(ocomenca.NumeOcom.Value);
                var distribucion = _detaivaService.ConvertirADistribucionDto(registros);
                
                _logger.LogInformation($"[GetDistribucionByOrden] Distribución creada - Total IVA: {distribucion.TotalIva:C2}, Total Base: {distribucion.TotalBaseImponible:C2}");
                
                return Ok(distribucion);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[GetDistribucionByOrden] Error al obtener distribución IVA para orden ID: {ocomencaId}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Borra los datos existentes de DETAIVA para la orden
        /// Equivalente a: DELETE DETAIVA where Nume_Ocom = numeroOrden
        /// Reutiliza el método existente EliminarDetalleIvaAsync del DetaivaService
        /// </summary>
        [HttpPost("borrar-por-orden")]
        public async Task<IActionResult> BorrarPorOrden([FromBody] BorrarDetaivaRequest request)
        {
            try
            {
                _logger.LogInformation($"[BorrarPorOrden] Eliminando distribución IVA para orden: {request.NumeroOrden}");

                // Usar el método existente del servicio
                var registrosEliminados = await _detaivaService.EliminarDetalleIvaAsync(request.NumeroOrden);

                _logger.LogInformation($"[BorrarPorOrden] Eliminados {registrosEliminados} registros DETAIVA");

                return Ok(new { 
                    Success = true, 
                    Message = $"DETAIVA eliminado exitosamente. Registros afectados: {registrosEliminados}",
                    RegistrosAfectados = registrosEliminados
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[BorrarPorOrden] Error al borrar DETAIVA para orden {request.NumeroOrden}");
                return StatusCode(500, new { 
                    Success = false, 
                    Message = "Error al borrar DETAIVA",
                    Error = ex.Message 
                });
            }
        }

        /// <summary>
        /// Inserta la distribución de IVA para una orden
        /// Reutiliza el método existente GrabarDetalleIvaAsync del DetaivaService
        /// </summary>
        [HttpPost("insertar-distribucion")]
        public async Task<IActionResult> InsertarDistribucion([FromBody] InsertarDistribucionRequest request)
        {
            try
            {
                _logger.LogInformation($"[InsertarDistribucion] Insertando distribución IVA para orden: {request.NumeroOrden}");
                _logger.LogInformation($"[InsertarDistribucion] Registros a insertar: {request.DistribucionIva.Count}");

                // Convertir la lista de objetos dinámicos a DistribucionIvaDTO
                var distribucionDto = ConvertirADistribucionDto(request.DistribucionIva);

                // Usar el método existente del servicio para grabar
                var resultado = await _detaivaService.GrabarDetalleIvaAsync(
                    request.NumeroOrden, 
                    0, // CodigoEmpresa - se asignará automáticamente por el AuditInterceptor
                    distribucionDto
                );

                if (resultado)
                {
                    _logger.LogInformation($"[InsertarDistribucion] ✅ Distribución IVA insertada exitosamente");
                    return Ok(new { 
                        Success = true, 
                        Message = $"Distribución IVA insertada exitosamente. Total registros: {request.DistribucionIva.Count}",
                        RegistrosInsertados = request.DistribucionIva.Count
                    });
                }
                else
                {
                    _logger.LogError($"[InsertarDistribucion] ❌ Error al insertar distribución IVA");
                    return StatusCode(500, new { 
                        Success = false, 
                        Message = "Error al insertar distribución IVA"
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[InsertarDistribucion] Error al insertar distribución IVA para orden {request.NumeroOrden}");
                return StatusCode(500, new { 
                    Success = false, 
                    Message = "Error al insertar distribución IVA",
                    Error = ex.Message 
                });
            }
        }

        /// <summary>
        /// Convierte la lista de objetos dinámicos a DistribucionIvaDTO
        /// </summary>
        private DistribucionIvaDTO ConvertirADistribucionDto(List<DistribucionIvaItem> items)
        {
            var distribucion = new DistribucionIvaDTO();

            foreach (var item in items)
            {
                switch (item.Norit)
                {
                    case 0: // 0%
                        distribucion.BaseImponible0 += item.BaseImpo;
                        break;
                    case 1: // 10.5%
                        distribucion.BaseImponible105 += item.BaseImpo;
                        distribucion.ImporteIva105 += item.ImpoIva;
                        break;
                    case 2: // 21%
                        distribucion.BaseImponible21 += item.BaseImpo;
                        distribucion.ImporteIva21 += item.ImpoIva;
                        break;
                    case 3: // 27%
                        distribucion.BaseImponible27 += item.BaseImpo;
                        distribucion.ImporteIva27 += item.ImpoIva;
                        break;
                }
            }

            return distribucion;
        }
    }

    /// <summary>
    /// DTO para la solicitud de borrar DETAIVA
    /// </summary>
    public class BorrarDetaivaRequest
    {
        public int NumeroOrden { get; set; }
        public string Condicion { get; set; } = string.Empty;
    }

    /// <summary>
    /// DTO para la solicitud de insertar distribución IVA
    /// </summary>
    public class InsertarDistribucionRequest
    {
        public int NumeroOrden { get; set; }
        public List<DistribucionIvaItem> DistribucionIva { get; set; } = new();
    }

    /// <summary>
    /// DTO para un item de distribución IVA
    /// </summary>
    public class DistribucionIvaItem
    {
        public int NumeOcom { get; set; }
        public int Norit { get; set; }
        public decimal TasaIva { get; set; }
        public decimal BaseImpo { get; set; }
        public decimal ImpoIva { get; set; }
    }
}
