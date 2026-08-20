using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/ubicasql")]
    public class UbicasqlController : ControllerBase
    {
        private readonly DataContext _context;

        public UbicasqlController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Devuelve las ubicaciones ACTIVAS (Status >= 0) para un producto.
        /// 
        /// Notas importantes:
        /// - En esta implementación, las ubicaciones no se borran físicamente; se usa baja lógica con Status = -1.
        /// - Por eso, este endpoint filtra (Status ?? 0) >= 0.
        /// - Se devuelve un DTO "liviano" para el front (Id + texto), ya trimmeado.
        /// </summary>
        [HttpGet("por-producto/{codint:int}")]
        public async Task<ActionResult<List<UbicacionProductoDTO>>> GetUbicacionesPorProductoAsync([FromRoute] int codint)
        {
            // Validación defensiva: si el código interno no es válido, no tiene sentido consultar la base.
            if (codint <= 0)
            {
                return Ok(new List<UbicacionProductoDTO>());
            }

            // Consulta EF:
            // - AsNoTracking(): son datos de lectura, no queremos overhead de tracking.
            // - Status>=0: solo ubicaciones activas.
            // - Trim(): normaliza el texto para el usuario.
            var ubicaciones = await _context.Ubicasqls
                .AsNoTracking()
                .Where(x => x.CodInte == codint && (x.Status ?? 0) >= 0)
                .Select(x => new UbicacionProductoDTO
                {
                    Id = x.Id,
                    Codiubica = (x.Codiubica ?? string.Empty).Trim()
                })
                // Seguridad extra: descartamos strings vacíos por si hubiera datos antiguos o inconsistentes.
                .Where(x => !string.IsNullOrWhiteSpace(x.Codiubica))
                .OrderBy(x => x.Codiubica)
                .ToListAsync();

            return Ok(ubicaciones);
        }

        /// <summary>
        /// Devuelve las ubicaciones DADAS DE BAJA (Status < 0) para un producto.
        /// 
        /// Esto se usa para que el usuario pueda restaurarlas (Status=0) sin recrearlas.
        /// </summary>
        [HttpGet("por-producto-baja/{codint:int}")]
        public async Task<ActionResult<List<UbicacionProductoDTO>>> GetUbicacionesBajaPorProductoAsync([FromRoute] int codint)
        {
            if (codint <= 0)
            {
                return Ok(new List<UbicacionProductoDTO>());
            }

            // Mismo criterio que el endpoint de activas, pero filtrando por Status<0.
            var ubicaciones = await _context.Ubicasqls
                .AsNoTracking()
                .Where(x => x.CodInte == codint && (x.Status ?? 0) < 0)
                .Select(x => new UbicacionProductoDTO
                {
                    Id = x.Id,
                    Codiubica = (x.Codiubica ?? string.Empty).Trim()
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.Codiubica))
                .OrderBy(x => x.Codiubica)
                .ToListAsync();

            return Ok(ubicaciones);
        }

        /// <summary>
        /// Crea una nueva ubicación para un producto.
        /// 
        /// Reglas de negocio implementadas acá (además de validaciones en front):
        /// - CodInte debe ser válido.
        /// - Codiubica es obligatorio, trim, máximo 32.
        /// - Status arranca en 0 (activa).
        /// - CodiEmpr se fuerza a 0 por convención del sistema.
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<UbicacionProductoDTO>> CrearUbicacionAsync([FromBody] UbicacionProductoCreateDTO dto)
        {
            if (dto == null || dto.CodInte <= 0)
            {
                return BadRequest("Código interno inválido.");
            }

            // Normalizamos el texto que llega del front: evitamos espacios al inicio/fin.
            var codiubica = (dto.Codiubica ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(codiubica))
            {
                return BadRequest("La ubicación es obligatoria.");
            }

            if (codiubica.Length > 32)
            {
                return BadRequest("La ubicación no puede superar 32 caracteres.");
            }

            // Persistimos:
            // - Status 0 = activa.
            // - CodiEmpr 0 = regla global del sistema.
            var entity = new Ubicasql
            {
                CodInte = dto.CodInte,
                CodiEmpr = 0,
                Codiubica = codiubica,
                Status = 0
            };

            _context.Ubicasqls.Add(entity);
            await _context.SaveChangesAsync();

            return Ok(new UbicacionProductoDTO
            {
                Id = entity.Id,
                Codiubica = entity.Codiubica ?? string.Empty
            });
        }

        /// <summary>
        /// Baja lógica de una ubicación (Status = -1).
        /// 
        /// Se usa para que el usuario pueda "eliminar" sin perder historial y permitiendo restauración.
        /// </summary>
        [HttpPut("baja/{id:int}")]
        public async Task<IActionResult> BajaUbicacionAsync([FromRoute] int id)
        {
            try
            {
                if (id <= 0)
                {
                    return BadRequest();
                }

                // No AsNoTracking: vamos a modificar la entidad.
                var ubica = await _context.Ubicasqls.FirstOrDefaultAsync(x => x.Id == id);
                if (ubica == null)
                {
                    return NotFound();
                }

                // Status negativo = dado de baja.
                ubica.Status = -1;

                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                // Aquí puedes registrar el error si tienes ILogger.
                // _logger.LogError(ex, "Error al dar de baja la ubicación con Id {Id}", id);

                return StatusCode(StatusCodes.Status500InternalServerError,
                    new
                    {
                        mensaje = "Ocurrió un error al dar de baja la ubicación.",
                        detalle = ex.Message
                    });
            }
        }

        /// <summary>
        /// Restauración de una ubicación dada de baja (Status = 0).
        /// </summary>
        [HttpPut("restaurar/{id:int}")]
        public async Task<IActionResult> RestaurarUbicacionAsync([FromRoute] int id)
        {
            if (id <= 0)
            {
                return BadRequest();
            }

            var ubica = await _context.Ubicasqls.FirstOrDefaultAsync(x => x.Id == id);
            if (ubica == null)
            {
                return NotFound();
            }

            // Status 0 = activa.
            ubica.Status = 0;
            await _context.SaveChangesAsync();

            return Ok();
        }

        /// <summary>
        /// Actualiza el texto (CODIUBICA) de una ubicación existente.
        /// 
        /// Reglas:
        /// - Trim, obligatorio, máximo 32.
        /// - No permitir duplicados para el mismo producto (CodInte), comparando case-insensitive.
        /// - Excluir el propio Id del chequeo de duplicado.
        /// </summary>
        [HttpPut("actualizar/{id:int}")]
        public async Task<IActionResult> ActualizarUbicacionAsync([FromRoute] int id, [FromBody] UbicacionProductoUpdateDTO dto)
        {
            if (id <= 0)
            {
                return BadRequest("ID inválido.");
            }

            if (dto == null)
            {
                return BadRequest("Datos inválidos.");
            }

            var texto = (dto.Codiubica ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(texto))
            {
                return BadRequest("La ubicación es obligatoria.");
            }

            if (texto.Length > 32)
            {
                return BadRequest("La ubicación no puede superar 32 caracteres.");
            }

            var ubica = await _context.Ubicasqls.FirstOrDefaultAsync(x => x.Id == id);
            if (ubica == null)
            {
                return NotFound();
            }

            // El CodInte (producto) se obtiene del registro actual.
            // Esto evita que el front tenga que enviar CodInte para editar.
            var codInte = ubica.CodInte ?? 0;
            if (codInte <= 0)
            {
                return BadRequest("Código interno inválido.");
            }

            // Normalizamos a MAYÚSCULAS para comparar "sin importar" may/min.
            var textoNormalized = texto.ToUpper();
            var existe = await _context.Ubicasqls
                .AsNoTracking()
                .Where(x => x.Id != id && x.CodInte == codInte)
                // Importante: la comparación se hace con Trim() y ToUpper() para que sea case-insensitive.
                .AnyAsync(x => ((x.Codiubica ?? string.Empty).Trim().ToUpper()) == textoNormalized);

            if (existe)
            {
                return BadRequest($"La ubicación '{texto}' ya existe para este código.");
            }

            // Finalmente persistimos el nuevo texto.
            ubica.Codiubica = texto;
            await _context.SaveChangesAsync();

            return Ok();
        }
    }
}
