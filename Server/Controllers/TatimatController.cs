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
    [Route("/api/tipomaterial")]
    public class TatimatController : ControllerBase
    {
        private readonly DataContext _context;

        public TatimatController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de tipos de material para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Tatimat
                .Where(d => d.Status >= 0)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de tipos de material con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado del tipo de material (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Tatimat.AsQueryable();

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
            var queryable = _context.Tatimat.AsQueryable();

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
        /// Obtiene un tipo de material por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var tatimat = await _context.Tatimat
                .FirstOrDefaultAsync(x => x.Id == id);
            if (tatimat == null)
            {
                return NotFound();
            }
            return Ok(tatimat);
        }

        /// <summary>
        /// Crea un nuevo tipo de material
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Tatimat tatimat)
        {
            try
            {
                // Verificar si ya existe un tipo de material con el mismo código
                var existingTatimat = await _context.Tatimat
                    .FirstOrDefaultAsync(x => x.CodigoTatimat == tatimat.CodigoTatimat);

                if (existingTatimat != null)
                {
                    return BadRequest($"Ya existe un tipo de material con el código {tatimat.CodigoTatimat}");
                }

                _context.Add(tatimat);
                await _context.SaveChangesAsync();
                return Ok(tatimat);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza un tipo de material existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Tatimat tatimat)
        {
            try
            {
                // Verificar si ya existe un tipo de material con el mismo código (excluyendo el actual)
                var existingTatimat = await _context.Tatimat
                    .FirstOrDefaultAsync(x => x.CodigoTatimat == tatimat.CodigoTatimat && x.Id != tatimat.Id);

                if (existingTatimat != null)
                {
                    return BadRequest($"Ya existe un tipo de material con el código {tatimat.CodigoTatimat}");
                }

                _context.Update(tatimat);
                await _context.SaveChangesAsync();
                return Ok(tatimat);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina un tipo de material (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var tatimat = await _context.Tatimat.FirstOrDefaultAsync(x => x.Id == id);
            if (tatimat == null)
            {
                return NotFound();
            }
            _context.Remove(tatimat);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca un tipo de material como inactivo (borrado lógico)
        /// </summary>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        {
            var tipomaterial = await _context.Tatimat.FirstOrDefaultAsync(x => x.Id == id);
            if (tipomaterial == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            tipomaterial.Status = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
