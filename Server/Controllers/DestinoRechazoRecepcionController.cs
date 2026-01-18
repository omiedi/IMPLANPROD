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
    [Route("/api/destinorechazorecepcion")]
    public class DestinoRechazoRecepcionController : ControllerBase
    {
        private readonly DataContext _context;

        public DestinoRechazoRecepcionController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de destinos de rechazo para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            try
            {
                var destinos = await _context.DestinoRechazoRecepciones
                    .Where(d => d.Status >= 0)
                    .Select(s => new { Key = s.CodigoDestino, Value = s.Descripcion })
                    .OrderBy(s => s.Value)
                    .ToDictionaryAsync(s => s.Key, s => s.Value);

                return new OkObjectResult(destinos);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Obtiene todos los destinos de rechazo activos
        /// </summary>
        [HttpGet("activos")]
        public async Task<ActionResult> GetActivos()
        {
            try
            {
                return new OkObjectResult(await _context.DestinoRechazoRecepciones
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
        /// Obtiene la lista paginada de destinos de rechazo con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado del destino (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.DestinoRechazoRecepciones.AsQueryable();
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
            catch (Exception exception)
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
            var queryable = _context.DestinoRechazoRecepciones.AsQueryable();
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
            catch (Exception exception)
            {
                throw;
            }

        }

        /// <summary>
        /// Obtiene un destino de rechazo por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var destino = await _context.DestinoRechazoRecepciones
                .FirstOrDefaultAsync(x => x.Id == id);
            if (destino == null)
            {
                return NotFound();
            }
            return new OkObjectResult(destino);
        }

        /// <summary>
        /// Crea un nuevo destino de rechazo
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(DestinoRechazoRecepcion destino)
        {
            try
            {
                // Verificar si ya existe un destino con el mismo código
                var existingDestino = await _context.DestinoRechazoRecepciones
                    .FirstOrDefaultAsync(x => x.CodigoDestino == destino.CodigoDestino);

                if (existingDestino != null)
                {
                    return BadRequest($"Ya existe un destino de rechazo con el código {destino.CodigoDestino}");
                }

                _context.Add(destino);
                await _context.SaveChangesAsync();
                return new OkObjectResult(destino);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza un destino de rechazo existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(DestinoRechazoRecepcion destino)
        {
            try
            {
                // Verificar si ya existe un destino con el mismo código (excluyendo el actual)
                var existingDestino = await _context.DestinoRechazoRecepciones
                    .FirstOrDefaultAsync(x => x.CodigoDestino == destino.CodigoDestino && x.Id != destino.Id);

                if (existingDestino != null)
                {
                    return BadRequest($"Ya existe un destino de rechazo con el código {destino.CodigoDestino}");
                }

                _context.Update(destino);
                await _context.SaveChangesAsync();
                return new OkObjectResult(destino);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina un destino de rechazo (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var destino = await _context.DestinoRechazoRecepciones.FirstOrDefaultAsync(x => x.Id == id);
            if (destino == null)
            {
                return NotFound();
            }
            _context.Remove(destino);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca un destino de rechazo como inactivo (borrado lógico)
        /// </summary>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusAsync(int id, int status = 0)
        {
            var destino = await _context.DestinoRechazoRecepciones.FirstOrDefaultAsync(x => x.Id == id);
            if (destino == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            destino.Status = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
