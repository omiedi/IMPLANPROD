using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Shared.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// EJEMPLO: Controlador que demuestra cómo funciona la auditoría automática
    /// NO se requiere código adicional para capturar información de auditoría
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class EjemploAuditoriaController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<EjemploAuditoriaController> _logger;
        
        // NOTA: Estos servicios NO se usan directamente en el controlador
        // El AuditInterceptor los usa automáticamente
        private readonly IUsuarioFlexoftService _usuarioFlexoftService;
        private readonly ICodiempresa _codiempresa;

        public EjemploAuditoriaController(
            DataContext context, 
            ILogger<EjemploAuditoriaController> logger,
            IUsuarioFlexoftService usuarioFlexoftService,
            ICodiempresa codiempresa)
        {
            _context = context;
            _logger = logger;
            _usuarioFlexoftService = usuarioFlexoftService;
            _codiempresa = codiempresa;
        }

        /// <summary>
        /// CREAR un nuevo movimiento de stock
        /// La auditoría se maneja AUTOMÁTICAMENTE
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<Movisto>> CrearMovimiento(Movisto movimiento)
        {
            try
            {
                // ============================================================
                // NO ES NECESARIO ESTABLECER CAMPOS DE AUDITORÍA MANUALMENTE
                // El AuditInterceptor lo hace automáticamente:
                // - movimiento.AddRecord = DateTime.Now
                // - movimiento.LastUpdate = DateTime.Now
                // - movimiento.NumUsuar = [ID del usuario actual]
                // - movimiento.CodiEmprNet = [Código de empresa del usuario]
                // ============================================================

                _context.Movistos.Add(movimiento);
                await _context.SaveChangesAsync();

                _logger.LogInformation($"Movimiento creado con ID: {movimiento.Id}");
                
                // En este punto, los campos de auditoría ya están establecidos automáticamente
                return CreatedAtAction(nameof(ObtenerMovimiento), new { id = movimiento.Id }, movimiento);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear movimiento");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// ACTUALIZAR un movimiento existente
        /// La auditoría se maneja AUTOMÁTICAMENTE
        /// </summary>
        [HttpPut("{id}")]
        public async Task<IActionResult> ActualizarMovimiento(int id, Movisto movimiento)
        {
            if (id != movimiento.Id)
                return BadRequest("ID no coincide");

            try
            {
                // ============================================================
                // NO ES NECESARIO ACTUALIZAR CAMPOS DE AUDITORÍA MANUALMENTE
                // El AuditInterceptor lo hace automáticamente:
                // - movimiento.LastUpdate = DateTime.Now (actualizado)
                // - movimiento.NumUsuar = [ID del usuario actual]
                // - movimiento.CodiEmprNet = [Código de empresa del usuario]
                // - movimiento.AddRecord = (se mantiene sin cambios)
                // ============================================================

                _context.Entry(movimiento).State = EntityState.Modified;
                await _context.SaveChangesAsync();

                _logger.LogInformation($"Movimiento {id} actualizado");
                
                return NoContent();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!MovimientoExists(id))
                    return NotFound();
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al actualizar movimiento {id}");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// OBTENER un movimiento con información de auditoría
        /// </summary>
        [HttpGet("{id}")]
        public async Task<ActionResult<Movisto>> ObtenerMovimiento(int id)
        {
            var movimiento = await _context.Movistos.FindAsync(id);

            if (movimiento == null)
                return NotFound();

            // Los campos de auditoría están disponibles automáticamente:
            // - movimiento.AddRecord (cuándo se creó)
            // - movimiento.LastUpdate (cuándo se modificó por última vez)
            // - movimiento.NumUsuar (quién lo creó/modificó)
            // - movimiento.CodiEmprNet (en qué empresa)

            return movimiento;
        }

        /// <summary>
        /// CONSULTAR auditoría de movimientos
        /// Ejemplo de cómo usar los campos de auditoría para reportes
        /// </summary>
        [HttpGet("auditoria")]
        public async Task<ActionResult<object>> ConsultarAuditoria(
            DateTime? fechaDesde = null,
            DateTime? fechaHasta = null,
            int? usuarioId = null,
            int? empresaId = null)
        {
            try
            {
                var query = _context.Movistos.AsQueryable();

                // Filtrar por fecha de creación
                if (fechaDesde.HasValue)
                    query = query.Where(m => m.AddRecord >= fechaDesde.Value);

                if (fechaHasta.HasValue)
                    query = query.Where(m => m.AddRecord <= fechaHasta.Value);

                // Filtrar por usuario
                if (usuarioId.HasValue)
                    query = query.Where(m => m.NumUsuar == usuarioId.Value);

                // Filtrar por empresa
                if (empresaId.HasValue)
                    query = query.Where(m => m.CodiEmprNet == empresaId.Value);

                var resultado = await query
                    .Select(m => new
                    {
                        m.Id,
                        m.Cod_Inte,
                        m.CantIngre,
                        m.CantSalid,
                        FechaCreacion = m.AddRecord,
                        UltimaModificacion = m.LastUpdate,
                        Usuario = m.NumUsuar,
                        Empresa = m.CodiEmprNet
                    })
                    .OrderByDescending(m => m.UltimaModificacion)
                    .Take(100)
                    .ToListAsync();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al consultar auditoría");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// EJEMPLO: Operación en lote que también se audita automáticamente
        /// </summary>
        [HttpPost("lote")]
        public async Task<ActionResult> CrearMovimientosEnLote(List<Movisto> movimientos)
        {
            try
            {
                // Todos los movimientos se auditan automáticamente
                _context.Movistos.AddRange(movimientos);
                await _context.SaveChangesAsync();

                // Cada movimiento tendrá automáticamente:
                // - AddRecord y LastUpdate establecidos
                // - NumUsuar y CodiEmprNet del usuario actual

                _logger.LogInformation($"Creados {movimientos.Count} movimientos en lote");
                return Ok($"Creados {movimientos.Count} movimientos");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear movimientos en lote");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// DEMOSTRACIÓN: Cómo obtener manualmente la información del usuario
        /// (Solo para referencia - NO es necesario en operaciones normales)
        /// </summary>
        [HttpGet("info-usuario")]
        public ActionResult<object> ObtenerInfoUsuario()
        {
            try
            {
                // Así es como el AuditInterceptor obtiene la información automáticamente:
                int? idFlexoft = _usuarioFlexoftService.InterObtenerIdFlexoft(User);
                short? codigoEmpresa = _codiempresa.IObtenerCodiempresa(User);

                return Ok(new
                {
                    UsuarioId = idFlexoft,
                    EmpresaId = codigoEmpresa,
                    Mensaje = "Esta información se captura automáticamente en cada operación"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener información del usuario");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        private bool MovimientoExists(int id)
        {
            return _context.Movistos.Any(e => e.Id == id);
        }
    }
}

/*
RESUMEN DEL EJEMPLO:

✅ AUTOMÁTICO - No se requiere código adicional para auditoría
✅ TRANSPARENTE - Los controladores funcionan normalmente
✅ COMPLETO - Captura usuario, empresa y timestamps automáticamente
✅ EFICIENTE - Sin impacto en la lógica de negocio existente

CAMPOS CAPTURADOS AUTOMÁTICAMENTE:
- AddRecord: Fecha de creación (solo en INSERT)
- LastUpdate: Fecha de última modificación (INSERT y UPDATE)
- NumUsuar: ID del usuario que realizó la operación
- CodiEmprNet: Código de empresa del usuario activo

OPERACIONES AUDITADAS:
- INSERT: Todos los campos de auditoría se establecen
- UPDATE: LastUpdate, NumUsuar y CodiEmprNet se actualizan
- DELETE: (Si se implementa soft delete, también se auditaría)

¡El sistema funciona completamente en segundo plano!
*/
