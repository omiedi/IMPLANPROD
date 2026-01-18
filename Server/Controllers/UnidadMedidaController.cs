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
    [Route("/api/umedidas")]
    public class UnidadMedidaController : ControllerBase
    {
        private readonly DataContext _context;

        public UnidadMedidaController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de umedidas para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.UnidadMedidas
                .Where(d => d.status >= 0)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de umedidas con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado de la umedida (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.UnidadMedidas.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion.ToLower().Contains(pagination.Filter.ToLower()));
            }

            // Filtrar por estado
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.status == estado || x.status == null); // Activos
            }
            else
            {
                queryable = queryable.Where(x => x.status == estado); // Inactivos
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
            var queryable = _context.UnidadMedidas.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion!.ToLower().Contains(pagination.Filter.ToLower()));
            }

            // Filtrar por estado
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.status == estado || x.status == null);
            }
            else
            {
                queryable = queryable.Where(x => x.status == estado);
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene una umedida por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var umdedida = await _context.UnidadMedidas
                .FirstOrDefaultAsync(x => x.Id == id);
            if (umdedida == null)
            {
                return NotFound();
            }
            return Ok(umdedida);
        }

        /// <summary>
        /// Crea una nueva umedida
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(UnidadMedidas unidadMedida)
        {
            try
            {
                // Verificar si ya existe una umedida con el mismo código
                var existingUmedida = await _context.UnidadMedidas
                    .FirstOrDefaultAsync(x => x.DescripcionAbreviada == unidadMedida.DescripcionAbreviada);

                if (existingUmedida != null)
                {
                    return BadRequest($"Ya existe una U.medida con el código {unidadMedida.DescripcionAbreviada}");
                }

                _context.Add(unidadMedida);
                await _context.SaveChangesAsync();
                return Ok(unidadMedida);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza una umedida existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(UnidadMedidas umedida)
        {
            try
            {
                // Verificar si ya existe una umedida con el mismo código (excluyendo la actual)
                var existingUmedida = await _context.UnidadMedidas
                    .FirstOrDefaultAsync(x => x.DescripcionAbreviada == umedida.DescripcionAbreviada && x.Id != umedida.Id);

                if (existingUmedida != null)
                {
                    return BadRequest($"Ya existe una U.medida con el código {umedida.DescripcionAbreviada}");
                }

                _context.Update(umedida);
                await _context.SaveChangesAsync();
                return Ok(umedida);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina una umedida (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var umedida = await _context.UnidadMedidas.FirstOrDefaultAsync(x => x.Id == id);
            if (umedida == null)
            {
                return NotFound();
            }
            _context.Remove(umedida);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca una umedida como inactiva (borrado lógico)
        /// </summary>
        [HttpPut("baja/{id:int}")]
        public async Task<IActionResult> umedidaStatusMenosUnoAsync(int id)
        {
            var umedida = await _context.UnidadMedidas.FirstOrDefaultAsync(x => x.Id == id);
            if (umedida == null)
            {
                return NotFound();
            }
            umedida.status = -1;
            _context.Update(umedida);
            await _context.SaveChangesAsync();
            return NoContent();
        }
        //______________________________________________________________________
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        {
            var umedida = await _context.UnidadMedidas.FirstOrDefaultAsync(x => x.Id == id);
            if (umedida == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            umedida.status = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
//modificar para que se modifique esta tabla
//ALTER TABLE UnidadMedidas
//ADD status INT NOT NULL;

//update UnidadMedidas set status = 0



