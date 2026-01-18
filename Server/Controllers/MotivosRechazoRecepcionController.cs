using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/motivosrechazorecepcion")]
    public class MotivosRechazoRecepcionController : ControllerBase
    {
        private readonly DataContext _context;

        public MotivosRechazoRecepcionController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de motivos de rechazo para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            try
            {
                var motivos = await _context.MotivosRechazoRecepciones
                    .Where(d => d.Status >= 0)
                    .Select(s => new { Key = s.CodigoMotivo, Value = s.Descripcion })
                    .OrderBy(s => s.Value)
                    .ToDictionaryAsync(s => s.Key, s => s.Value);

                return Ok(motivos);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Obtiene todos los motivos de rechazo activos
        /// </summary>
        [HttpGet("activos")]
        public async Task<ActionResult> GetActivos()
        {
            try
            {
                return Ok(await _context.MotivosRechazoRecepciones
                    .Where(d => d.Status >= 0)
                    .OrderBy(s => s.Descripcion)
                    .ToListAsync());
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Obtiene la lista paginada de motivos de rechazo con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado del motivo (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.MotivosRechazoRecepciones.AsQueryable();
            try
            {
                // Aplicar filtro si existe
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    queryable = queryable.Where(x => x.Descripcion!.ToLower().Contains(pagination.Filter.ToLower()));
                }

                // Filtrar por estado
                if (estado == 0)
                {
                    queryable = queryable.Where(x => x.Status >= 0); // Activos
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == estado); // Inactivos
                }

                return Ok(await queryable
                    .OrderBy(x => x.Descripcion)
                    .Paginate(pagination)
                    .ToListAsync());
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.MotivosRechazoRecepciones.AsQueryable();
            try
            {
                // Aplicar filtro si existe
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    queryable = queryable.Where(x => x.Descripcion!.ToLower().Contains(pagination.Filter.ToLower()));
                }

                // Filtrar por estado
                if (estado == 0)
                {
                    queryable = queryable.Where(x => x.Status >= 0); // Activos
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == estado); // Inactivos
                }

                double count = await queryable.CountAsync();
                double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
                return Ok(totalPages);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Obtiene un motivo de rechazo por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var motivo = await _context.MotivosRechazoRecepciones
                .FirstOrDefaultAsync(x => x.Id == id);
            if (motivo == null)
            {
                return NotFound();
            }
            return Ok(motivo);
        }

        /// <summary>
        /// Crea un nuevo motivo de rechazo
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(MotivosRechazoRecepcion motivo)
        {
            try
            {
                // Verificar si ya existe un motivo con el mismo código
                var existingMotivo = await _context.MotivosRechazoRecepciones
                    .FirstOrDefaultAsync(x => x.CodigoMotivo == motivo.CodigoMotivo);

                if (existingMotivo != null)
                {
                    return BadRequest($"Ya existe un motivo de rechazo con el código {motivo.CodigoMotivo}");
                }

                _context.Add(motivo);
                await _context.SaveChangesAsync();
                return Ok(motivo);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza un motivo de rechazo existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(MotivosRechazoRecepcion motivo)
        {
            try
            {
                // Verificar si ya existe un motivo con el mismo código (excluyendo el actual)
                var existingMotivo = await _context.MotivosRechazoRecepciones
                    .FirstOrDefaultAsync(x => x.CodigoMotivo == motivo.CodigoMotivo && x.Id != motivo.Id);

                if (existingMotivo != null)
                {
                    return BadRequest($"Ya existe un motivo de rechazo con el código {motivo.CodigoMotivo}");
                }

                _context.Update(motivo);
                await _context.SaveChangesAsync();
                return Ok(motivo);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina un motivo de rechazo (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var motivo = await _context.MotivosRechazoRecepciones.FirstOrDefaultAsync(x => x.Id == id);
            if (motivo == null)
            {
                return NotFound();
            }
            _context.Remove(motivo);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca un motivo de rechazo como inactivo (borrado lógico)
        /// </summary>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusAsync(int id, int status = 0)
        {
            var motivo = await _context.MotivosRechazoRecepciones.FirstOrDefaultAsync(x => x.Id == id);
            if (motivo == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            motivo.Status = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
