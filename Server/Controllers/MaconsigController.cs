using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class MaconsigController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<MaconsigController> _logger;
        private readonly ConsignacionService _consignacionService;

        public MaconsigController(DataContext context, ILogger<MaconsigController> logger, ConsignacionService consignacionService)
        {
            _context = context;
            _logger = logger;
            _consignacionService = consignacionService;
        }

        /// <summary>
        /// Obtiene materiales en consignación con saldo pendiente
        /// Filtro: Cant_Rendid < Cant_Entreg - 0.001
        /// </summary>
        [HttpGet("materiales-consignacion")]
        public async Task<ActionResult<List<Maconsig>>> GetMaterialesConsignacionAsync()
        {
            try
            {
                _logger.LogInformation("[GetMaterialesConsignacion] Obteniendo materiales en consignación");

                var materiales = await _context.Maconsigs
                    .Where(m => m.CantRendid < m.CantEntreg - 0.001m)
                    .OrderByDescending(m => m.FemiOcom)
                    .ThenBy(m => m.NroOcom)
                    .ThenBy(m => m.CodInte)
                    .ToListAsync();

                _logger.LogInformation($"[GetMaterialesConsignacion] Se encontraron {materiales.Count} materiales en consignación");

                return Ok(materiales);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[GetMaterialesConsignacion] Error al obtener materiales en consignación");
                return StatusCode(500, "Error al obtener materiales en consignación");
            }
        }

        /// <summary>
        /// Borra consignaciones sin entregas de MACONSIG
        /// Equivalente a: DELETE FROM MACONSIG WHERE NRO_OCOM LIKE 'OC-000001%' AND Cant_Entreg = 0
        /// </summary>
        [HttpPost("borrar-consignaciones")]
        public async Task<IActionResult> BorrarConsignaciones([FromBody] BorrarConsignacionesRequest request)
        {
            try
            {
                _logger.LogInformation($"[BorrarConsignaciones] Procesando orden: {request.NumeroOrden}");
                _logger.LogInformation($"[BorrarConsignaciones] Condición: {request.Condicion}");

                // Construir la consulta SQL para borrar consignaciones
                var sql = $"DELETE FROM MACONSIG WHERE NRO_OCOM LIKE '{request.CodigoFormateado}%' AND Cant_Entreg = 0";
                
                _logger.LogInformation($"[BorrarConsignaciones] Ejecutando SQL: {sql}");

                // Ejecutar la consulta
                var registrosAfectados = await _context.Database.ExecuteSqlRawAsync(sql);

                _logger.LogInformation($"[BorrarConsignaciones] Registros eliminados: {registrosAfectados}");

                return Ok(new { 
                    Success = true, 
                    Message = $"Consignaciones eliminadas exitosamente. Registros afectados: {registrosAfectados}",
                    RegistrosAfectados = registrosAfectados
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[BorrarConsignaciones] Error al borrar consignaciones para orden {request.NumeroOrden}");
                
                // Si no hay registros para borrar, no es un error
                if (ex.Message.Contains("no rows affected") || ex.Message.Contains("0 registros"))
                {
                    return Ok(new { 
                        Success = true, 
                        Message = "No hay consignaciones para borrar",
                        RegistrosAfectados = 0
                    });
                }

                return StatusCode(500, new { 
                    Success = false, 
                    Message = "Ha ocurrido un error inesperado, tener en cuenta que puede ser que no haya consignaciones para dicha orden de compra",
                    Error = ex.Message 
                });
            }
        }

        /// <summary>
        /// Procesa las consignaciones para una orden de compra según la lógica VB6
        /// Implementa la lógica de remanente: CANTIASIGN# = CANTIASIGN# - CANYASIG#
        /// 
        /// USO: TANTO EN ALTA COMO EN EDICIÓN DE ÓRDENES DE COMPRA
        /// - ALTA: Se ejecuta desde OcomencaController.PostAsync() dentro de la transacción
        /// - EDICIÓN: Se ejecuta en el PASO 6 del flujo de edición (OrdenesCompraEditForm.razor)
        /// - Calcula remanente considerando consignaciones ya existentes
        /// - Solo crea registros si remanente > 0.05 y componente tiene CodXElem
        /// </summary>
        [HttpPost("procesar-orden")]
        public async Task<IActionResult> ProcesarConsignacionesOrdenAsync([FromBody] ProcesarConsignacionesRequest request)
        {
            try
            {
                _logger.LogInformation($"[ProcesarConsignaciones] 🔄 Procesando consignaciones para O.C. {request.NumeroOrden}");
                _logger.LogInformation($"[ProcesarConsignaciones] Proveedor: {request.NumeroProveedor} - {request.RazonSocialProveedor}");
                _logger.LogInformation($"[ProcesarConsignaciones] Fecha: {request.FechaEmision:yyyy-MM-dd}, Moneda: {request.MonedaEmision}");

                // Usar transacción para garantizar integridad
                using var transaction = await _context.Database.BeginTransactionAsync();
                try
                {
                    var resultado = await _consignacionService.ProcesarConsignacionesOrdenCompraAsync(
                        request.NumeroOrden,
                        request.NumeroProveedor,
                        request.RazonSocialProveedor,
                        request.FechaEmision,
                        request.MonedaEmision
                    );

                    // ===== GUARDAR CAMBIOS EN LA BASE DE DATOS =====
                    if (resultado)
                    {
                        await _context.SaveChangesAsync();
                        await transaction.CommitAsync();
                        _logger.LogInformation($"[ProcesarConsignaciones] 💾 Consignaciones guardadas y confirmadas en BD para O.C. {request.NumeroOrden}");
                    }
                    else
                    {
                        await transaction.CommitAsync(); // Commit aunque no haya consignaciones
                    }

                    var response = new
                    {
                        Success = resultado,
                        Message = resultado 
                            ? "Consignaciones procesadas y guardadas exitosamente"
                            : "No se encontraron items con estructura para procesar",
                        NumeroOrden = request.NumeroOrden
                    };

                    _logger.LogInformation($"[ProcesarConsignaciones] ✅ Resultado: {response.Message}");
                    return Ok(response);
                }
                catch (Exception)
                {
                    await transaction.RollbackAsync();
                    throw; // Re-lanzar para que lo capture el catch externo
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[ProcesarConsignaciones] ❌ Error procesando consignaciones para O.C. {request.NumeroOrden}");
                return StatusCode(500, new { 
                    Success = false,
                    Message = $"Error al procesar consignaciones: {ex.Message}",
                    NumeroOrden = request.NumeroOrden
                });
            }
        }
    }

    /// <summary>
    /// DTO para la solicitud de borrar consignaciones
    /// USO: SOLO EN EDICIÓN - PASO 1 del flujo de edición
    /// </summary>
    public class BorrarConsignacionesRequest
    {
        public int NumeroOrden { get; set; }
        public string CodigoFormateado { get; set; } = string.Empty;
        public string Condicion { get; set; } = string.Empty;
    }

    /// <summary>
    /// DTO para la solicitud de procesamiento de consignaciones
    /// USO: TANTO EN ALTA COMO EN EDICIÓN
    /// - ALTA: Desde OcomencaController.PostAsync()
    /// - EDICIÓN: PASO 6 del flujo de edición
    /// </summary>
    public class ProcesarConsignacionesRequest
    {
        public int NumeroOrden { get; set; }
        public int NumeroProveedor { get; set; }
        public string RazonSocialProveedor { get; set; } = string.Empty;
        public DateTime FechaEmision { get; set; }
        public short MonedaEmision { get; set; }
    }
}
