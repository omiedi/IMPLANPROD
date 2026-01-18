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
    /// Controlador para la gestión de vendedores
    /// </summary>
    [ApiController]
    [Route("/api/vendedores")]
    public class VendedorController : ControllerBase
    {
        private readonly DataContext _context;

        public VendedorController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de vendedores para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Vendedores
                .Where(d => d.Status >= 0)
                .OrderBy(x => x.NombreVendedor)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de vendedores con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado del vendedor (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Vendedores.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.NombreVendedor!.ToLower().Contains(pagination.Filter.ToLower()) ||
                                                x.NumVendedor.ToString().Contains(pagination.Filter));
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
                .OrderBy(x => x.NombreVendedor)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Vendedores.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.NombreVendedor!.ToLower().Contains(pagination.Filter.ToLower()) ||
                                                x.NumVendedor.ToString().Contains(pagination.Filter));
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
        /// Obtiene un vendedor por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var vendedor = await _context.Vendedores
                .FirstOrDefaultAsync(x => x.Id == id);
            if (vendedor == null)
            {
                return NotFound();
            }
            return Ok(vendedor);
        }

        /// <summary>
        /// Crea un nuevo vendedor
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Vendedor vendedor)
        {
            try
            {
                // Verificar si ya existe un vendedor con el mismo número
                if (vendedor.NumVendedor.HasValue)
                {
                    var existingVendedor = await _context.Vendedores
                        .FirstOrDefaultAsync(x => x.NumVendedor == vendedor.NumVendedor);

                    if (existingVendedor != null)
                    {
                        return BadRequest($"Ya existe un vendedor con el número {vendedor.NumVendedor}");
                    }
                }

                // Establecer status activo para nuevos vendedores
                vendedor.Status = 0;

                _context.Add(vendedor);
                await _context.SaveChangesAsync();
                return Ok(vendedor);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza un vendedor existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Vendedor vendedor)
        {
            try
            {
                // Verificar si ya existe un vendedor con el mismo número (excluyendo el actual)
                if (vendedor.NumVendedor.HasValue)
                {
                    var existingVendedor = await _context.Vendedores
                        .FirstOrDefaultAsync(x => x.NumVendedor == vendedor.NumVendedor && x.Id != vendedor.Id);

                    if (existingVendedor != null)
                    {
                        return BadRequest($"Ya existe un vendedor con el número {vendedor.NumVendedor}");
                    }
                }

                _context.Update(vendedor);
                await _context.SaveChangesAsync();
                return Ok(vendedor);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina un vendedor (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var vendedor = await _context.Vendedores.FirstOrDefaultAsync(x => x.Id == id);
            if (vendedor == null)
            {
                return NotFound();
            }
            _context.Remove(vendedor);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca un vendedor como inactivo (borrado lógico)
        /// </summary>
        [HttpPut("baja/{id:int}")]
        public async Task<IActionResult> VendedorStatusMenosUnoAsync(int id)
        {
            var vendedor = await _context.Vendedores.FirstOrDefaultAsync(x => x.Id == id);
            if (vendedor == null)
            {
                return NotFound();
            }
            vendedor.Status = -1;
            _context.Update(vendedor);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Cambia el estado de un vendedor
        /// </summary>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        {
            var vendedor = await _context.Vendedores.FirstOrDefaultAsync(x => x.Id == id);
            if (vendedor == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            vendedor.Status = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
