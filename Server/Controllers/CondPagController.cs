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
    /// Controlador para la gestión de condiciones de pago
    /// </summary>
    [ApiController]
    [Route("/api/condpag")]
    public class CondPagController : ControllerBase
    {
        private readonly DataContext _context;

        public CondPagController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de condiciones de pago para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.CondPags
                .Where(d => d.NumCondPag >= 0)
                .OrderBy(x => x.DescCondPag)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de condiciones de pago con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado de la condición de pago (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.CondPags.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.DescCondPag.ToLower().Contains(pagination.Filter.ToLower()) ||
                                                x.NumCondPag.ToString().Contains(pagination.Filter));
            }

            // Filtrar por estado
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.NumCondPag >= 0 || x.NumCondPag == null); // Activos
            }
            else
            {
                queryable = queryable.Where(x => x.NumCondPag == estado); // Inactivos
            }

            return Ok(await queryable
                .OrderBy(x => x.DescCondPag)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.CondPags.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.DescCondPag!.ToLower().Contains(pagination.Filter.ToLower()) ||
                                                x.NumCondPag.ToString().Contains(pagination.Filter));
            }

            // Filtrar por estado
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.NumCondPag >= 0 || x.NumCondPag == null);
            }
            else
            {
                queryable = queryable.Where(x => x.NumCondPag == estado);
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene una condición de pago por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var condPag = await _context.CondPags
                .FirstOrDefaultAsync(x => x.Id == id);
            if (condPag == null)
            {
                return NotFound();
            }
            return Ok(condPag);
        }

        // En el controlador de condiciones de pago
        [HttpGet("condiciones-pago/{condPago}")]
        public async Task<ActionResult<CondicionPagoDTO>> GetCondicionPago(short condPago)
        {
            var condPag = await _context.CondPags
                .FirstOrDefaultAsync(x => x.NumCondPag == condPago);
            if (condPag == null)
            {
                return NotFound();
            }
            return Ok(condPag);
        }

        /// <summary>
        /// Crea una nueva condición de pago
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(CondPag condPag)
        {
            try
            {
                // Verificar si ya existe una condición de pago con el mismo número
                if (condPag.NumCondPag.HasValue)
                {
                    var existingCondPag = await _context.CondPags
                        .FirstOrDefaultAsync(x => x.NumCondPag == condPag.NumCondPag);

                    if (existingCondPag != null)
                    {
                        return BadRequest($"Ya existe una condición de pago con el número {condPag.NumCondPag}");
                    }
                }

                _context.Add(condPag);
                await _context.SaveChangesAsync();
                return Ok(condPag);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza una condición de pago existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(CondPag condPag)
        {
            try
            {
                // Verificar si ya existe una condición de pago con el mismo número (excluyendo la actual)
                if (condPag.NumCondPag.HasValue)
                {
                    var existingCondPag = await _context.CondPags
                        .FirstOrDefaultAsync(x => x.NumCondPag == condPag.NumCondPag && x.Id != condPag.Id);

                    if (existingCondPag != null)
                    {
                        return BadRequest($"Ya existe una condición de pago con el número {condPag.NumCondPag}");
                    }
                }

                _context.Update(condPag);
                await _context.SaveChangesAsync();
                return Ok(condPag);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina una condición de pago (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var condPag = await _context.CondPags.FirstOrDefaultAsync(x => x.Id == id);
            if (condPag == null)
            {
                return NotFound();
            }
            _context.Remove(condPag);
            await _context.SaveChangesAsync();
            return NoContent();
        }


        /// <summary>
        /// Cambia el estado de una condición de pago
        /// </summary>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        {
            var condPag = await _context.CondPags.FirstOrDefaultAsync(x => x.Id == id);
            if (condPag == null)
            {
                return NotFound();
            }

            // Cambiar el campo NumCondPag como indicador de estado
            condPag.NumCondPag = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
