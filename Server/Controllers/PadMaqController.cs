using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using IMPLANPROD.Server.Helpers;
using Microsoft.EntityFrameworkCore;


namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/padmaq")]
    public class PadMaqController : ControllerBase
    {
        private readonly DataContext _context;

        public PadMaqController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de maquina para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.padmaqs
                .Where(d => d.Status >= 0)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de maquina con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado del tipo de material (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.padmaqs.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion.ToLower().Contains(pagination.Filter.ToLower()));
            }

            // Filtrar por estado
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.Status == estado || x.Status == null); // Activos
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

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.padmaqs.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion.ToLower().Contains(pagination.Filter.ToLower()));
            }

            // Filtrar por estado
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
        /// Obtiene un maquina por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var padmaq = await _context.padmaqs
                .FirstOrDefaultAsync(x => x.Id == id);
            if (padmaq == null)
            {
                return NotFound();
            }
            return Ok(padmaq);
        }

        /// <summary>
        /// Crea un nuevo maquina
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(PADMAQ padmaq)
        {
            try
            {
                // Verificar si ya existe maquina con el mismo código
                var existingpadmaq = await _context.padmaqs
                    .FirstOrDefaultAsync(x => x.CodigoPadMaq == padmaq.CodigoPadMaq);

                if (existingpadmaq != null)
                {
                    return BadRequest($"Ya existe un máquina con el código {padmaq.CodigoPadMaq}");
                }

                _context.Add(padmaq);
                await _context.SaveChangesAsync();
                return Ok(padmaq);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza un maquina existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(PADMAQ padmaq)
        {
            try
            {
                // Verificar si ya existe un tipo de material con el mismo código (excluyendo el actual)
                var existingpadmaq = await _context.padmaqs
                    .FirstOrDefaultAsync(x => x.CodigoPadMaq == padmaq.CodigoPadMaq && x.Id != padmaq.Id);

                if (existingpadmaq != null)
                {
                    return BadRequest($"Ya existe una máquina con el código {padmaq.CodigoPadMaq}");
                }

                _context.Update(padmaq);
                await _context.SaveChangesAsync();
                return Ok(padmaq);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina un maquina (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var padmaq = await _context.padmaqs.FirstOrDefaultAsync(x => x.Id == id);
            if (padmaq == null)
            {
                return NotFound();
            }
            _context.Remove(padmaq);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca una maquina como inactivo (borrado lógico)
        /// </summary>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        {
            var padmaq = await _context.padmaqs.FirstOrDefaultAsync(x => x.Id == id);
            if (padmaq == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            padmaq.Status = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
