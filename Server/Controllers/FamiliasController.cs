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
    [Route("/api/familias")]
    public class FamiliasController : ControllerBase
    {
        private readonly DataContext _context;

        public FamiliasController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de familias para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Familias
                .Where(d => d.Status >= 0)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de familias con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado de la familia (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Familias.AsQueryable();

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
            var queryable = _context.Familias.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion!.ToLower().Contains(pagination.Filter.ToLower()));
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
        /// Obtiene una familia por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var familia = await _context.Familias
                .FirstOrDefaultAsync(x => x.Id == id);
            if (familia == null)
            {
                return NotFound();
            }
            return Ok(familia);
        }

        /// <summary>
        /// Crea una nueva familia
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Familia familia)
        {
            try
            {
                // Verificar si ya existe una familia con el mismo código
                var existingFamilia = await _context.Familias
                    .FirstOrDefaultAsync(x => x.CodigoFamilia == familia.CodigoFamilia);

                if (existingFamilia != null)
                {
                    return BadRequest($"Ya existe una familia con el código {familia.CodigoFamilia}");
                }

                _context.Add(familia);
                await _context.SaveChangesAsync();
                return Ok(familia);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza una familia existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Familia familia)
        {
            try
            {
                // Verificar si ya existe una familia con el mismo código (excluyendo la actual)
                var existingFamilia = await _context.Familias
                    .FirstOrDefaultAsync(x => x.CodigoFamilia == familia.CodigoFamilia && x.Id != familia.Id);

                if (existingFamilia != null)
                {
                    return BadRequest($"Ya existe una familia con el código {familia.CodigoFamilia}");
                }

                _context.Update(familia);
                await _context.SaveChangesAsync();
                return Ok(familia);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina una familia (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var familia = await _context.Familias.FirstOrDefaultAsync(x => x.Id == id);
            if (familia == null)
            {
                return NotFound();
            }
            _context.Remove(familia);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca una familia como inactiva (borrado lógico)
        /// </summary>
        [HttpPut("baja/{id:int}")]
        public async Task<IActionResult> FamiliaStatusMenosUnoAsync(int id)
        {
            var familia = await _context.Familias.FirstOrDefaultAsync(x => x.Id == id);
            if (familia == null)
            {
                return NotFound();
            }
            familia.Status = -1;
            _context.Update(familia);
            await _context.SaveChangesAsync();
            return NoContent();
        }
        //______________________________________________________________________
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        {
            var familia = await _context.Familias.FirstOrDefaultAsync(x => x.Id == id);
            if (familia == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            familia.Status = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
