using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para gestión de Transportes.
    /// Sigue el patrón estándar de controladores del proyecto.
    /// </summary>
    [ApiController]
    [Route("/api/transportes")]
    public class TransportesController : ControllerBase
    {
        private readonly DataContext _context;

        public TransportesController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Devuelve lista para combos (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Transportes
                .Where(t => t.Status >= 0 || t.Status == null)
                .OrderBy(t => t.Deno_Tra)
                .ToListAsync());
        }

        /// <summary>
        /// Lista paginada con filtros
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Transportes.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                var f = pagination.Filter.ToLower();
                queryable = queryable.Where(x =>
                    (x.Deno_Tra ?? string.Empty).ToLower().Contains(f) ||
                    (x.Raso_Tra ?? string.Empty).ToLower().Contains(f) ||
                    (x.Cuit_Tra ?? string.Empty).ToLower().Contains(f));
            }

            if (estado == 0)
            {
                queryable = queryable.Where(x => x.Status == estado || x.Status == null);
            }
            else
            {
                queryable = queryable.Where(x => x.Status == estado);
            }

            return Ok(await queryable
                .OrderBy(x => x.Deno_Tra)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Total de páginas para paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Transportes.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                var f = pagination.Filter.ToLower();
                queryable = queryable.Where(x =>
                    (x.Deno_Tra ?? string.Empty).ToLower().Contains(f) ||
                    (x.Raso_Tra ?? string.Empty).ToLower().Contains(f) ||
                    (x.Cuit_Tra ?? string.Empty).ToLower().Contains(f));
            }

            if (estado == 0)
            {
                queryable = queryable.Where(x => x.Status == estado || x.Status == null);
            }
            else
            {
                queryable = queryable.Where(x => x.Status == estado);
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene un transporte por Id
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var item = await _context.Transportes.FirstOrDefaultAsync(x => x.Id == id);
            if (item == null) return NotFound();
            return Ok(item);
        }

        /// <summary>
        /// Crea un nuevo transporte
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Transporte transporte)
        {
            try
            {
                _context.Add(transporte);
                await _context.SaveChangesAsync();
                return Ok(transporte);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza un transporte existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Transporte transporte)
        {
            try
            {
                _context.Update(transporte);
                await _context.SaveChangesAsync();
                return Ok(transporte);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Borrado físico por Id
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var item = await _context.Transportes.FirstOrDefaultAsync(x => x.Id == id);
            if (item == null) return NotFound();
            _context.Remove(item);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Cambio de estado (baja lógica o reactivación)
        /// </summary>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> CambiaStatusAsync(int id, int status = 0)
        {
            var item = await _context.Transportes.FirstOrDefaultAsync(x => x.Id == id);
            if (item == null) return NotFound();
            item.Status = (short)status;
            await _context.SaveChangesAsync();
            return NoContent();
        }
    }
}
