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
    [Route("/api/secprod")]
    public class SecProdController : ControllerBase
    {
        private readonly DataContext _context;

        public SecProdController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de sectores productivos para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            try
            {
                var sectores = await _context.SecProds
                    .Where(d => d.Status >= 0)
                    .Select(s => new { Key = s.CodigoSecProd, Value = s.Descripcion })
                    .OrderBy(s => s.Value)
                    .ToDictionaryAsync(s => s.Key, s => s.Value);

                return new OkObjectResult(sectores);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Obtiene todos los sectores productivos activos
        /// </summary>
        [HttpGet("activos")]
        public async Task<ActionResult> GetActivos()
        {
            try
            {
                return new OkObjectResult(await _context.SecProds
                    .Where(d => d.Status >= 0)
                    .OrderBy(s => s.Descripcion)
                    .ToListAsync());
            }
            catch (Exception)
            {
                throw;
            }
        }
        [HttpGet("combo2")]

        public async Task<ActionResult> GetCombo2()
        {
            return new OkObjectResult(_context.SecProds.ToListAsync());
        }
        /// <summary>
        /// Obtiene la lista paginada de sector productivo con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado del tipo de material (0=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.SecProds.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Descripcion!.ToLower().Contains(pagination.Filter.ToLower()));
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

            //return Ok(await queryable
            //    .OrderBy(x => x.Descripcion)
            //    .Paginate(pagination)
            //    .ToListAsync());
            return new OkObjectResult(await queryable
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
            var queryable = _context.SecProds.AsQueryable();

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
            //return Ok(totalPages);
            return new OkObjectResult(totalPages);
        }

        /// <summary>
        /// Obtiene un sector productivo por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var secprod = await _context.SecProds
                .FirstOrDefaultAsync(x => x.Id == id);
            if (secprod == null)
            {
                return NotFound();
            }
            //return Ok(secprod);
            return new OkObjectResult(secprod);
        }

        //private IActionResult Ok(object secprod)
        //{
        //    throw new NotImplementedException();
        //}

        /// <summary>
        /// Crea un nuevo sector productivo
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(SecProd secprod)
        {
            try
            {
                // Verificar si ya existe un sector productivo con el mismo código
                var existingsecprod = await _context.SecProds
                    .FirstOrDefaultAsync(x => x.CodigoSecProd == secprod.CodigoSecProd);

                if (existingsecprod != null)
                {
                    return BadRequest($"Ya existe un sector productivo con el código {secprod.CodigoSecProd}");
                }

                _context.Add(secprod);
                await _context.SaveChangesAsync();
                //return (ActionResult)Ok(secprod);
                return new OkObjectResult(secprod);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza un sector productivo existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(SecProd secprod)
        {
            try
            {
                // Verificar si ya existe un sector productivo con el mismo código (excluyendo el actual)
                var existingsecprod = await _context.SecProds
                    .FirstOrDefaultAsync(x => x.CodigoSecProd == secprod.CodigoSecProd && x.Id != secprod.Id);

                if (existingsecprod != null)
                {
                    return BadRequest($"Ya existe un sector productivo con el código {secprod.CodigoSecProd}");
                }

                _context.Update(secprod);
                await _context.SaveChangesAsync();
                //return Ok(secprod);
                return new OkObjectResult(secprod);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina un sector productivo (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var secprod = await _context.SecProds.FirstOrDefaultAsync(x => x.Id == id);
            if (secprod == null)
            {
                return NotFound();
            }
            _context.Remove(secprod);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca un sector productivo como inactivo (borrado lógico)
        /// </summary>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        {
            var secprod = await _context.SecProds.FirstOrDefaultAsync(x => x.Id == id);
            if (secprod == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            secprod.Status = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
        /// <summary>
        /// Retorna los depósitos asociados a una lista de códigos de sector productivo.
        /// </summary>
        /// <param name="codigos">Lista de códigos separados por coma (ej: 100,200,300)</param>
        /// <returns>Lista de depósitos únicos</returns>
        [HttpGet("depositos")]
        public async Task<ActionResult<List<int>>> ObtenerDepositosPorCodigos([FromQuery] string codigos)
        {
            if (string.IsNullOrWhiteSpace(codigos))
            {
                return BadRequest(new { mensaje = "Debe proporcionar al menos un código de sector." });
            }

            var listaCodigos = codigos
                .Split(',')
                .Select(c => int.TryParse(c, out int val) ? val : (int?)null)
                .Where(c => c.HasValue)
                .Select(c => c.Value)
                .ToList();

            if (!listaCodigos.Any())
            {
                return BadRequest(new { mensaje = "Los códigos proporcionados no son válidos." });
            }

            var depositos = await _context.SecProds
                .Where(s => listaCodigos.Contains(s.CodigoSecProd))
                .Select(s => s.Deposito)
                .Where(d => d.HasValue)
                .Select(d => d.Value)
                .Distinct()
                .ToListAsync();

            return new ActionResult<List<int>>(depositos);
        }

        [AllowAnonymous]
        [HttpGet("comboList")]
        public async Task<ActionResult<List<SecProd>>> GetComboList()
        {
            var sectores = await _context.SecProds
                .AsNoTracking()
                .Where(d => d.Status >= 0)
                .OrderBy(s => s.Descripcion)
                .ToListAsync();

            return Ok(sectores);
        }



    }
}
