using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
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

        /// <summary>
        /// Recupera los datos necesarios para imprimir el "Plan de Inspección y Ensayo" en formato A4.
        /// Combina el producto (Master), el encabezado del plan (Codiproc) y el detalle (Plainsp).
        /// Se usa desde el botón "Imprimir" del modal PlanInspeccionModal (/ingenieria/documentos).
        /// GET /api/codiproc/{codInte}/datos-impresion-plan-inspeccion
        /// </summary>
        /// <param name="codInte">Código interno del producto (Master.Codint / Codiproc.Cod_Inte).</param>
        /// <param name="origen">
        /// Origen opcional del plan (por ejemplo "OF-000123-1") cuando se imprime desde el
        /// contexto de una Orden de Fabricación. Si no se envía, queda vacío (impresión
        /// desde /ingenieria/documentos, donde el plan no está atado a ninguna OF puntual).
        /// </param>
        [HttpGet("{codInte:int}/datos-impresion-plan-inspeccion")]
        public async Task<IActionResult> ObtenerDatosImpresionPlanInspeccion(int codInte, [FromQuery] string? origen = null)
        {
            // Datos del producto (encabezado: Código, Denominación, Revisión).
            var master = await _context.masters
                .AsNoTracking()
                .Where(m => m.Codint == codInte)
                .FirstOrDefaultAsync();

            if (master == null)
            {
                return NotFound("No se encontró el producto solicitado.");
            }

            // Encabezado del plan de inspección (AQL, Nivel, Frecuencia, Puntos de Detención, Preparó/Aprobó).
            var plan = await _context.Codiprocs
                .AsNoTracking()
                .Where(x => x.Cod_Inte == codInte)
                .FirstOrDefaultAsync();

            // Detalle de PLAINSP ordenado por NroOrden, tal como lo definió el usuario en el modal.
            // No se filtran filas vacías: pueden ser separadores visuales intencionales entre bloques.
            var items = await _context.Plainsps
                .AsNoTracking()
                .Where(x => x.Cod_Inte == codInte)
                .OrderBy(x => x.NroOrden)
                .ThenBy(x => x.Id)
                .Select(x => new PlanInspeccionItemImpresionDTO
                {
                    DimSegunPlano = x.DimSegunPlano,
                    Instrumento = x.Instrumento,
                    Observaciones = x.Observaciones
                })
                .ToListAsync();

            var dto = new PlanInspeccionImpresionDTO
            {
                CodInte = master.Codint,
                Codigo = master.Codigo,
                Denominacion = master.Descripcion,
                Revision = master.Revision,
                // Cuando se imprime desde la Orden de Fabricación se recibe el IdSubOr (p.ej.
                // "OF-000123-1") por query string. Desde /ingenieria/documentos no se envía
                // y queda vacío, porque el plan no está atado a ninguna OF puntual.
                Origen = origen ?? string.Empty,
                FechaImpresionFormateada = DateTime.Now.ToString("dd/MM/yyyy"),
                AQL = plan?.AQL,
                NivelPct = plan?.NivelPct,
                FrecInspeccion = plan?.FrecInspeccion,
                PuntoDetencion = plan?.PuntoDetencion,
                NombrePreparo = plan?.NombrePreparo,
                FechaPreparoFormateada = plan?.FechaPreparo?.ToString("dd/MM/yyyy") ?? string.Empty,
                NombreAprobo = plan?.NombreAprobo,
                FechaAproboFormateada = plan?.FechaAprobo?.ToString("dd/MM/yyyy") ?? string.Empty,
                Observaciones = plan?.Observaciones,
                Items = items
            };

            return Ok(dto);
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
