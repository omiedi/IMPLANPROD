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
    [Route("/api/tamoneds")]
    public class TamonedsController : ControllerBase
    {
        private readonly DataContext _context;

        public TamonedsController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de monedas para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Tamoneds
                .Where(d => d.Status >= 0)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de monedas con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado de la moneda (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Tamoneds.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion.ToLower().Contains(pagination.Filter.ToLower()) ||
                                               x.CodigoMoneda.ToString().Contains(pagination.Filter));
            }

            // Filtrar por estado
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.Status == 0);
            }

            return Ok(await queryable
                .OrderBy(x => x.CodigoMoneda)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Tamoneds.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion.ToLower().Contains(pagination.Filter.ToLower()) ||
                                               x.CodigoMoneda.ToString().Contains(pagination.Filter));
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene una moneda por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<ActionResult> GetAsync(int id)
        {
            var tamoned = await _context.Tamoneds.FirstOrDefaultAsync(x => x.Id == id);
            if (tamoned == null)
            {
                return NotFound();
            }
            return Ok(tamoned);
        }

        /// <summary>
        /// Crea una nueva moneda
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Tamoned tamoned)
        {
            try
            {
                _context.Add(tamoned);
                await _context.SaveChangesAsync();
                return Ok(tamoned);
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
        /// Actualiza una moneda existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Tamoned tamoned)
        {
            try
            {
                _context.Update(tamoned);
                await _context.SaveChangesAsync();
                return Ok(tamoned);
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
        /// Elimina una moneda
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var tamoned = await _context.Tamoneds.FirstOrDefaultAsync(x => x.Id == id);
            if (tamoned == null)
            {
                return NotFound();
            }

            _context.Remove(tamoned);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Cambia el estado de una moneda
        /// </summary>
        [HttpPut("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> CambiaStatusAsync(int id, int status)
        {
            var tamoned = await _context.Tamoneds.FirstOrDefaultAsync(x => x.Id == id);
            if (tamoned == null)
            {
                return NotFound();
            }

            tamoned.Status = status;
            _context.Update(tamoned);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Obtiene todas las monedas inactivas
        /// </summary>
        [HttpGet("GetAllInactivos")]
        public async Task<ActionResult> GetAllInactivos()
        {
            return Ok(await _context.Tamoneds
                .Where(x => x.Status < 0)
                .OrderBy(x => x.CodigoMoneda)
                .ToListAsync());
        }
    }
}
