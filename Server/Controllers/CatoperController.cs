using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/catoper")]
    public class CatoperController : ControllerBase
    {
        private readonly DataContext _context;

        public CatoperController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de categorias de operario para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Catopers
                .Where(d => d.Status >= 0)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de categorias de operario con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado de la moneda (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Catopers.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion.ToLower().Contains(pagination.Filter.ToLower()) ||
                                               x.CodigoCategoria.ToString().Contains(pagination.Filter));
            }

            // Filtrar por estado
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.Status == 0);
            }

            return Ok(await queryable
                .OrderBy(x => x.CodigoCategoria)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Catopers.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion.ToLower().Contains(pagination.Filter.ToLower()) ||
                                               x.CodigoCategoria.ToString().Contains(pagination.Filter));
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene una categorias de operario por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<ActionResult> GetAsync(int id)
        {
            var tamoned = await _context.Catopers.FirstOrDefaultAsync(x => x.Id == id);
            if (tamoned == null)
            {
                return NotFound();
            }
            return Ok(tamoned);
        }

        /// <summary>
        /// Crea una nueva categorias de operario
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Catoper catoper)
        {
            try
            {
                var existiCat = await _context.Catopers
                  .FirstOrDefaultAsync(x => x.CodigoCategoria == catoper.CodigoCategoria);

                if (existiCat != null)
                {
                    return BadRequest($"Ya existe un máquina con el código {catoper.CodigoCategoria}");
                }
                _context.Add(catoper);
                await _context.SaveChangesAsync();
                return Ok(catoper);
            }
            catch (DbUpdateException dbUpdateException)
            {
                if (dbUpdateException.InnerException!.Message.Contains("duplicate"))
                {
                    return BadRequest("Ya existe una moneda con el mismo código.");
                }
                return BadRequest(dbUpdateException.Message);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza una categorias de operario existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Catoper catoper)
        {
            try
            {
                _context.Update(catoper);
                await _context.SaveChangesAsync();
                return Ok(catoper);
            }
            catch (DbUpdateException dbUpdateException)
            {
                if (dbUpdateException.InnerException!.Message.Contains("duplicate"))
                {
                    return BadRequest("Ya existe una moneda con el mismo código.");
                }
                return BadRequest(dbUpdateException.Message);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina una categorias de operario
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var catoper = await _context.Catopers.FirstOrDefaultAsync(x => x.Id == id);
            if (catoper == null)
            {
                return NotFound();
            }

            _context.Remove(catoper);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Cambia el estado de una categorias de operario
        /// </summary>
        [HttpPut("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> CambiaStatusAsync(int id, int status)
        {
            var categoria = await _context.Catopers.FirstOrDefaultAsync(x => x.Id == id);
            if (categoria == null)
            {
                return NotFound();
            }

            categoria.Status = status;
            _context.Update(categoria);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Obtiene todas las categorias de operario inactivas
        /// </summary>
        [HttpGet("GetAllInactivos")]
        public async Task<ActionResult> GetAllInactivos()
        {
            return Ok(await _context.Catopers
                .Where(x => x.Status < 0)
                .OrderBy(x => x.CodigoCategoria)
                .ToListAsync());
        }
    }
}
