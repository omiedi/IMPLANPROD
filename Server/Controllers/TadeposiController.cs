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
    [Route("api/tadeposi")]
    public class TadeposiController : ControllerBase
    {
        private readonly DataContext _context;

        public TadeposiController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de depósitos para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult<Dictionary<int, string>>> GetCombo()
        {
            try
            {
                var depositos = await _context.Tadeposi
                    .Where(d => d.Status >= 0)
                    .ToDictionaryAsync(
                        d => d.CodigoDepo,
                        d => $"{d.CodigoDepo} - {d.Descripcion}"
                    );

                return Ok(depositos);
            }
            catch (Exception ex)
            {
                return BadRequest("Error al obtener la lista de depósitos");
            }
        }

        /// <summary>
        /// Obtiene la lista paginada de depósitos con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado del depósito (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Tadeposi.AsQueryable();

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
            var queryable = _context.Tadeposi.AsQueryable();

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
        /// Obtiene un depósito por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var deposito = await _context.Tadeposi
                .FirstOrDefaultAsync(x => x.Id == id);
            if (deposito == null)
            {
                return NotFound();
            }
            return Ok(deposito);
        }
        /// <summary>
        /// Obtiene la descripcion del depósito por su ID
        /// </summary>
       
            ///
        [HttpGet("descripcion/{CodigoDepo}")]
        public async Task<ActionResult<string>> GetDescripcion(int CodigoDepo)
        {
            try
            {
                var tadeposi = await _context.Tadeposi
                .FirstOrDefaultAsync(m => m.CodigoDepo == CodigoDepo);

                if (tadeposi == null)
                    return Ok(string.Empty); //  Devuelve un ActionResult<string>

                string descripcion = tadeposi.Descripcion?.Trim() ?? string.Empty;
                return Ok(descripcion); //  Envuelto en Ok(...)
            }
            catch (Exception)
            {
                return BadRequest($"No se Encontro la descripción del depósito: {CodigoDepo}");
            }           
        }
        
        /// <summary>
        /// Crea un nuevo depósito
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Tadeposi deposito)
        {
            try
            {
                // Verificar si ya existe un depósito con el mismo código
                var existingDeposito = await _context.Tadeposi
                    .FirstOrDefaultAsync(x => x.CodigoDepo == deposito.CodigoDepo);

                if (existingDeposito != null)
                {
                    return BadRequest($"Ya existe un depósito con el código {deposito.CodigoDepo}");
                }

                _context.Add(deposito);
                await _context.SaveChangesAsync();
                return Ok(deposito);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza un depósito existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Tadeposi deposito)
        {
            try
            {
                // Verificar si ya existe un depósito con el mismo código (excluyendo el actual)
                var existingDeposito = await _context.Tadeposi
                    .FirstOrDefaultAsync(x => x.CodigoDepo == deposito.CodigoDepo && x.Id != deposito.Id);

                if (existingDeposito != null)
                {
                    return BadRequest($"Ya existe un depósito con el código {deposito.CodigoDepo}");
                }

                _context.Update(deposito);
                await _context.SaveChangesAsync();
                return Ok(deposito);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina un depósito (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var deposito = await _context.Tadeposi.FirstOrDefaultAsync(x => x.Id == id);
            if (deposito == null)
            {
                return NotFound();
            }
            _context.Remove(deposito);
            await _context.SaveChangesAsync();
            return NoContent();
        }
        [HttpGet("SoloCodigos")]
        public async Task<ActionResult<List<int>>> GetSoloCodigos()
        {
            try
            {
                var depositos = await _context.Tadeposi
                    .Select(d => d.CodigoDepo)
                    .Distinct()
                    .OrderBy(d => d)
                    .ToListAsync();

                return Ok(depositos);
            }
            catch (Exception ex)
            {
                return BadRequest("Error al obtener la lista de depósitos");
            }
        }
        ///// <summary>
        ///// Marca un depósito como inactivo (borrado lógico)
        ///// </summary>
        //[HttpPut("baja/{id:int}")]
        //public async Task<IActionResult> DepositoStatusMenosUnoAsync(int id)
        //{
        //    var deposito = await _context.Tadeposis.FirstOrDefaultAsync(x => x.Id == id);
        //    if (deposito == null)
        //    {
        //        return NotFound();
        //    }
        //    deposito.Status = -1;
        //    _context.Update(deposito);
        //    await _context.SaveChangesAsync();
        //    return NoContent();
        //}
        //______________________________________________________________________
        // <summary>
        /// Marca un depósito como inactivo (borrado lógico)
        /// </summary>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> cambiaStatus(int id, int status = 0)
        {
            var deposito = await _context.Tadeposi.FirstOrDefaultAsync(x => x.Id == id);
            if (deposito == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            deposito.Status =status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }

        /// <summary>
        /// Obtiene la lista de depósitos activos ordenados por código
        /// </summary>
        [AllowAnonymous]
        [HttpGet("activos")]
        public async Task<ActionResult<List<Tadeposi>>> GetDepositosActivos()
        {
            try
            {
                var depositos = await _context.Tadeposi
                    .Where(d => d.Status >= 0)
                    .OrderBy(d => d.CodigoDepo)
                    .ToListAsync();

                return Ok(depositos);
            }
            catch (Exception ex)
            {
                return BadRequest("Error al obtener la lista de depósitos activos");
            }
        }
    }
}
