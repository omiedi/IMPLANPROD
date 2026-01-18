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
    [Route("/api/grupivas")]
    public class GrupivaController : ControllerBase
    {
        private readonly DataContext _context;

        public GrupivaController(DataContext context)
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
            return Ok(await _context.Grupivas
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
            var queryable = _context.Grupivas.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Referencia.ToLower().Contains(pagination.Filter.ToLower()) ||
                                               x.CodigoID.ToString().Contains(pagination.Filter));
            }

            // Filtrar por estado
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.Status == 0);
            }

            return Ok(await queryable
                .OrderBy(x => x.CodigoID)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Grupivas.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Referencia.ToLower().Contains(pagination.Filter.ToLower()) ||
                                               x.CodigoID.ToString().Contains(pagination.Filter));
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
            var grupiva = await _context.Grupivas.FirstOrDefaultAsync(x => x.Id == id);
            if (grupiva == null)
            {
                return NotFound();
            }
            return Ok(grupiva);
        }

        /// <summary>
        /// Crea una nueva moneda
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Grupiva grupiva)
        {
            try
            {
                _context.Add(grupiva);
                await _context.SaveChangesAsync();
                return Ok(grupiva);
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
        /// Actualiza una grupo de iva existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Grupiva grupiva)
        {
            try
            {
                _context.Update(grupiva);
                await _context.SaveChangesAsync();
                return Ok(grupiva);
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
        /// Elimina una grupo de iva
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var grupiva = await _context.Grupivas.FirstOrDefaultAsync(x => x.Id == id);
            if (grupiva == null)
            {
                return NotFound();
            }

            _context.Remove(grupiva);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Cambia el estado de una grupo de iva
        /// </summary>
        [HttpPut("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> CambiaStatusAsync(int id, int status)
        {
            var grupiva = await _context.Grupivas.FirstOrDefaultAsync(x => x.Id == id);
            if (grupiva == null)
            {
                return NotFound();
            }

            grupiva.Status = status;
            _context.Update(grupiva);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Obtiene todas las monedas inactivas
        /// </summary>
        [HttpGet("GetAllInactivos")]
        public async Task<ActionResult> GetAllInactivos()
        {
            return Ok(await _context.Grupivas
                .Where(x => x.Status < 0)
                .OrderBy(x => x.CodigoID)
                .ToListAsync());
        }
    }
}