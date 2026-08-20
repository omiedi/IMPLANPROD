using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la tabla CODIPROC.
    /// Expone los datos del proceso de inspección de un producto:
    /// AQL, nivel, frecuencia, puntos de detención, firma de preparación/aprobación
    /// y observaciones generales.
    /// Tabla origen VB6: CODIPROC.DAT (512 bytes por registro)
    /// </summary>
    [ApiController]
    [Route("api/codiproc")]
    public class CodiprocController : ControllerBase
    {
        private readonly DataContext _context;

        public CodiprocController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Devuelve el registro CODIPROC asociado a un producto.
        /// Se usa en el encabezado del modal "Plan de Inspección y Ensayo"
        /// para mostrar AQL, Nivel, Frecuencia, Puntos de Detención,
        /// Preparó, Aprobó y Observaciones.
        /// GET /api/codiproc/by-codinte/{codInte}
        /// </summary>
        /// <param name="codInte">Código interno del producto</param>
        [HttpGet("by-codinte/{codInte:int}")]
        public async Task<ActionResult<Codiproc>> GetByCodInte(int codInte)
        {
            var item = await _context.Codiprocs
                .AsNoTracking()                     // solo lectura
                .Where(x => x.Cod_Inte == codInte)
                .FirstOrDefaultAsync();

            if (item == null) return NotFound();
            return item;
        }

        /// <summary>
        /// Actualiza los campos editables del encabezado del Plan de Inspección.
        /// Se invoca al presionar "Guardar" en el modal, junto con el guardado
        /// de las filas PLAINSP.
        /// Campos actualizados:
        ///   - NombrePreparo / FechaPreparo : firma del usuario que preparó
        ///   - NombreAprobo  / FechaAprobo  : firma del usuario que aprobó
        ///   - Observaciones                : texto libre (hasta 512 chars)
        /// PUT /api/codiproc/observaciones/{codInte}
        /// </summary>
        /// <param name="codInte">Código interno del producto a actualizar</param>
        /// <param name="dto">Campos editables del plan de inspección</param>
        [HttpPut("observaciones/{codInte:int}")]
        public async Task<ActionResult> UpdateObservaciones(int codInte, [FromBody] CodiprocDetalleDto dto)
        {
            // Buscar con tracking para poder modificar y guardar
            var entity = await _context.Codiprocs
                .FirstOrDefaultAsync(x => x.Cod_Inte == codInte);

            if (entity == null) return NotFound();

            // Actualizar todos los campos editables desde el modal
            entity.AQL           = dto.AQL;
            entity.NivelPct      = dto.NivelPct;
            entity.FrecInspeccion  = dto.FrecInspeccion;
            entity.PuntoDetencion  = dto.PuntoDetencion;
            entity.NombrePreparo = dto.NombrePreparo;
            entity.FechaPreparo  = dto.FechaPreparo;
            entity.NombreAprobo  = dto.NombreAprobo;
            entity.FechaAprobo   = dto.FechaAprobo;
            entity.Observaciones = dto.Observaciones;
            entity.LastUpdate    = DateTime.Now;     // auditoría: momento de la última modificación

            await _context.SaveChangesAsync();
            return Ok(new { updated = true });
        }
    }

    /// <summary>
    /// DTO para la actualización de CODIPROC desde el modal "Plan de Inspección y Ensayo".
    /// Incluye todos los campos editables del encabezado del plan.
    /// </summary>
    public record CodiprocDetalleDto(
        int?      AQL,             // Nivel de Calidad Aceptable (int)
        int?      NivelPct,        // Nivel porcentual de inspección (int)
        string?   FrecInspeccion,  // Frecuencia de Inspección (máx 32 chars)
        string?   PuntoDetencion,  // Puntos de Detención Obligatoria (máx 48 chars)
        string?   NombrePreparo,   // Nombre de quien preparó el plan (máx 24 chars)
        DateTime? FechaPreparo,    // Fecha de preparación
        string?   NombreAprobo,    // Nombre de quien aprobó el plan (máx 24 chars)
        DateTime? FechaAprobo,     // Fecha de aprobación
        string?   Observaciones);  // Observaciones generales (máx 512 chars)
}
