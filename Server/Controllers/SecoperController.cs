using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/secoper")]
    public class SecoperController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ISyncService _syncService;
        private readonly ILogger<SecoperController> _logger;

        public SecoperController(DataContext context, ISyncService syncService, ILogger<SecoperController> logger)
        {
            _context = context;
            _syncService = syncService;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene la lista paginada de operaciones con filtros
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Secopers
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => 
                    (x.CodXConj != null && x.CodXConj.ToLower().Contains(pagination.Filter.ToLower())) ||
                    (x.DescOper != null && x.DescOper.ToLower().Contains(pagination.Filter.ToLower())));
            }

            return Ok(await queryable
                .OrderBy(x => x.CodIConj)
                .ThenBy(x => x.NumeOper)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene todas las operaciones de un producto
        /// </summary>
        [HttpGet("producto/{codigoProducto:int}")]
        public async Task<IActionResult> GetOperacionesProductoAsync(int codigoProducto)
        {
            try
            {
                var operaciones = await _context.Secopers
                    .Where(s => s.CodIConj == codigoProducto)
                    .OrderBy(s => s.NumeOper)
                    .ToListAsync();

                return Ok(operaciones);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener las operaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("pages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Secopers.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x =>
                    (x.CodXConj != null && x.CodXConj.ToLower().Contains(pagination.Filter.ToLower())) ||
                    (x.DescOper != null && x.DescOper.ToLower().Contains(pagination.Filter.ToLower())));
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene el ID de la última operación creada para un producto y número de operación
        /// </summary>
        [HttpGet("ultimoid/{codiConj:int}/{numeOper:int}")]
        public async Task<ActionResult<int>> GetUltimoIdAsync(int codiConj, int numeOper)
        {
            try
            {
                var operacion = await _context.Secopers
                    .Where(x => x.CodIConj == codiConj && x.NumeOper == numeOper)
                    .OrderByDescending(x => x.Id)
                    .FirstOrDefaultAsync();

                if (operacion == null)
                {
                    return NotFound($"No se encontró la operación para el producto {codiConj} y número {numeOper}");
                }

                return Ok(operacion.Id);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// Obtiene una operación por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var operacion = await _context.Secopers
                .FirstOrDefaultAsync(x => x.Id == id);

            if (operacion == null)
            {
                return NotFound();
            }

            return Ok(operacion);
        }

        /// <summary>
        /// Guarda o actualiza las operaciones de un producto
        /// </summary>
        [HttpPost("producto/{codigoProducto:int}")]
        public async Task<ActionResult<List<Secoper>>> SaveOperacionesProductoAsync(int codigoProducto, [FromBody] List<Secoper> operaciones)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();

            try
            {
                // Eliminar operaciones existentes
                var operacionesExistentes = await _context.Secopers
                    .Where(s => s.CodIConj == codigoProducto)
                    .ToListAsync();

                _context.Secopers.RemoveRange(operacionesExistentes);

                // Lista para almacenar las operaciones guardadas
                var operacionesGuardadas = new List<Secoper>();

                // Agregar nuevas operaciones
                foreach (var operacion in operaciones)
                {
                    operacion.Id = 0; // Asegurar que se cree como nuevo
                    operacion.CodIConj = codigoProducto;
                    var entidad = _context.Secopers.Add(operacion);
                    operacionesGuardadas.Add(entidad.Entity); // Agregar la entidad con su ID temporal
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                // Generar archivo de sincronización si está habilitado
                try
                {
                    await _syncService.GenerarArchivoSyncSecOperAsync(codigoProducto);
                }
                catch (Exception syncEx)
                {
                    _logger.LogError(syncEx, $"Error al generar archivo de sincronización para SecOper CodIConj={codigoProducto}");
                }

                // Retornar la lista de operaciones con sus IDs asignados
                return Ok(operacionesGuardadas);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return BadRequest($"Error al guardar las operaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Crea una nueva operación
        /// </summary>
        /// 
        [HttpPost]
        public async Task<ActionResult<Secoper>> PostAsync(Secoper operacion)
        {
            try
            {
                // Verificar si ya existe una operación con el mismo CodIConj y NumeOper
                var existente = await _context.Secopers
                    .AnyAsync(x => x.CodIConj == operacion.CodIConj && x.NumeOper == operacion.NumeOper);

                if (existente)
                {
                    return BadRequest($"Ya existe una operación con el número {operacion.NumeOper} para el producto {operacion.CodIConj}");
                }

                _context.Add(operacion);
                await _context.SaveChangesAsync();

                // Devolver la operación completa con su ID asignado
                return Ok(operacion);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        //[HttpPost]
        //public async Task<ActionResult<Secoper>> PostAsync(Secoper operacion)
        //{
        //    try
        //    {
        //        // Verificar si ya existe una operación con el mismo CodIConj y NumeOper
        //        var existente = await _context.Secopers
        //            .AnyAsync(x => x.CodIConj == operacion.CodIConj && x.NumeOper == operacion.NumeOper);

        //        if (existente)
        //        {
        //            return BadRequest($"Ya existe una operación con el número {operacion.NumeOper} para el producto {operacion.CodIConj}");
        //        }

        //        _context.Add(operacion);
        //        await _context.SaveChangesAsync();

        //        // Devolver la operación completa con su ID asignado
        //        return Ok(operacion);
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }
        //}

        /// <summary>
        /// Actualiza una operación existente por su ID
        /// </summary>
        /// <param name="id">ID de la operación a actualizar</param>
        /// <param name="operacion">Datos actualizados de la operación</param>
        /// <remarks>
        /// Se utiliza [FromBody] para indicar explícitamente que el objeto Secoper debe ser deserializado
        /// desde el cuerpo (body) de la petición HTTP. Esto es especialmente importante en peticiones PUT
        /// donde enviamos un objeto complejo en formato JSON desde el cliente. Sin este atributo, el modelo
        /// podría no enlazarse correctamente ya que ASP.NET Core no sabría de dónde obtener los datos.
        /// </remarks>
        [HttpPut("{id:int}")]
        public async Task<ActionResult<Secoper>> PutAsync(int id, Secoper operacion)
        {
            try
            {
                if (operacion.Id != id)
                {
                    return BadRequest("El ID de la operación no coincide");
                }

                _context.Update(operacion);
                await _context.SaveChangesAsync();
                return Ok(operacion);
            }
            catch (Exception exception)
            {
                return BadRequest($"Error al actualizar la operación: {exception.Message}");
            }
        }

        /// <summary>
        /// Elimina una operación
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var operacion = await _context.Secopers.FirstOrDefaultAsync(x => x.Id == id);
            if (operacion == null)
            {
                return NotFound();
            }

            _context.Remove(operacion);
            await _context.SaveChangesAsync();
            return NoContent();
        }
    }
}
