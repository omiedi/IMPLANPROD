using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using IMPLANPROD.Server.Helpers;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para gestionar las asignaciones de usuarios a sectores productivos
    /// Implementa operaciones CRUD para la relación muchos a muchos Usuario-SecProd
    /// </summary>
    [ApiController]
    [Route("/api/usuariossectorprod")]
    public class UsuariosSectorProdController : ControllerBase
    {
        private readonly DataContext _context;

        /// <summary>
        /// Constructor del controlador
        /// </summary>
        /// <param name="context">Contexto de datos de Entity Framework</param>
        public UsuariosSectorProdController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de asignaciones usuario-sector productivo para combo (solo activos)
        /// </summary>
        /// <returns>Lista de asignaciones activas</returns>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            try
            {
                return new OkObjectResult(await _context.UsuariosSectorProds
                    .Where(d => d.Status >= 0)
                    .Select(x => new { x.Id, Descripcion = $"{x.NombreUsuario} - {x.Descripcion}" })
                    .ToListAsync());
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Obtiene todas las asignaciones activas
        /// </summary>
        /// <returns>Lista completa de asignaciones activas</returns>
        [HttpGet("activos")]
        public async Task<ActionResult> GetActivos()
        {
            try
            {
                return new OkObjectResult(await _context.UsuariosSectorProds
                    .Where(d => d.Status >= 0)
                    .OrderBy(x => x.NombreUsuario)
                    .ThenBy(x => x.Descripcion)
                    .ToListAsync());
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Obtiene la lista paginada de asignaciones usuario-sector productivo con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado de la asignación (0=activo, -1=inactivo)</param>
        /// <returns>Lista paginada de asignaciones</returns>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.UsuariosSectorProds.AsQueryable();

            // Aplicar filtro si existe (busca en nombre de usuario o descripción del sector)
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => 
                    x.NombreUsuario!.ToLower().Contains(pagination.Filter.ToLower()) ||
                    x.Descripcion!.ToLower().Contains(pagination.Filter.ToLower()));
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

            return new OkObjectResult(await queryable
                .OrderBy(x => x.NombreUsuario)
                .ThenBy(x => x.Descripcion)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado de la asignación (0=activo, -1=inactivo)</param>
        /// <returns>Número total de páginas</returns>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.UsuariosSectorProds.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => 
                    x.NombreUsuario!.ToLower().Contains(pagination.Filter.ToLower()) ||
                    x.Descripcion!.ToLower().Contains(pagination.Filter.ToLower()));
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
            return new OkObjectResult(totalPages);
        }

        /// <summary>
        /// Obtiene una asignación usuario-sector productivo por su ID
        /// </summary>
        /// <param name="id">ID de la asignación</param>
        /// <returns>Asignación encontrada o NotFound</returns>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var usuarioSectorProd = await _context.UsuariosSectorProds
                .FirstOrDefaultAsync(x => x.Id == id);
            
            if (usuarioSectorProd == null)
            {
                return NotFound();
            }
            
            return new OkObjectResult(usuarioSectorProd);
        }

        /// <summary>
        /// Verifica si ya existe una asignación para un usuario y sector productivo específicos
        /// </summary>
        /// <param name="idFlexoft">ID Flexoft del usuario</param>
        /// <param name="codigoSecProd">Código del sector productivo</param>
        /// <param name="excludeId">ID a excluir de la búsqueda (para edición)</param>
        /// <returns>True si existe, False si no existe</returns>
        [HttpGet("existe/{idFlexoft:int}/{codigoSecProd:int}/{excludeId:int?}")]
        public async Task<ActionResult<bool>> ExisteAsignacion(int idFlexoft, int codigoSecProd, int? excludeId = null)
        {
            var query = _context.UsuariosSectorProds
                .Where(x => x.IdFlexoft == idFlexoft && 
                           x.CodigoSecProd == codigoSecProd && 
                           x.Status >= 0);

            if (excludeId.HasValue)
            {
                query = query.Where(x => x.Id != excludeId.Value);
            }

            var existe = await query.AnyAsync();
            return new OkObjectResult(existe);
        }

        /// <summary>
        /// Obtiene las asignaciones de un usuario específico
        /// </summary>
        /// <param name="idFlexoft">ID Flexoft del usuario</param>
        /// <returns>Lista de asignaciones del usuario</returns>
        [HttpGet("usuario/{idFlexoft:int}")]
        public async Task<ActionResult> GetByUsuario(int idFlexoft)
        {
            var asignaciones = await _context.UsuariosSectorProds
                .Where(x => x.IdFlexoft == idFlexoft && x.Status >= 0)
                .OrderBy(x => x.Descripcion)
                .ToListAsync();

            return new OkObjectResult(asignaciones);
        }

        /// <summary>
        /// Obtiene las asignaciones de un sector productivo específico
        /// </summary>
        /// <param name="codigoSecProd">Código del sector productivo</param>
        /// <returns>Lista de asignaciones del sector</returns>
        [HttpGet("sector/{codigoSecProd:int}")]
        public async Task<ActionResult> GetBySector(int codigoSecProd)
        {
            var asignaciones = await _context.UsuariosSectorProds
                .Where(x => x.CodigoSecProd == codigoSecProd && x.Status >= 0)
                .OrderBy(x => x.NombreUsuario)
                .ToListAsync();

            return new OkObjectResult(asignaciones);
        }

        /// <summary>
        /// Crea una nueva asignación usuario-sector productivo
        /// </summary>
        /// <param name="usuarioSectorProd">Datos de la nueva asignación</param>
        /// <returns>Asignación creada o error</returns>
        [HttpPost]
        public async Task<ActionResult> PostAsync(UsuariosSectorProd usuarioSectorProd)
        {
            try
            {
                // Verificar si ya existe una asignación para este usuario y sector
                var existingAsignacion = await _context.UsuariosSectorProds
                    .FirstOrDefaultAsync(x => x.IdFlexoft == usuarioSectorProd.IdFlexoft && 
                                            x.CodigoSecProd == usuarioSectorProd.CodigoSecProd &&
                                            x.Status >= 0);

                if (existingAsignacion != null)
                {
                    return BadRequest($"Ya existe una asignación para el usuario {usuarioSectorProd.NombreUsuario} y el sector {usuarioSectorProd.Descripcion}");
                }

                // Establecer fecha de creación
                usuarioSectorProd.FechaCreacion = DateTime.Now;
                usuarioSectorProd.Status = 0; // Activo

                _context.Add(usuarioSectorProd);
                await _context.SaveChangesAsync();
                return new OkObjectResult(usuarioSectorProd);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Actualiza una asignación usuario-sector productivo existente
        /// </summary>
        /// <param name="usuarioSectorProd">Datos actualizados de la asignación</param>
        /// <returns>Asignación actualizada o error</returns>
        [HttpPut]
        public async Task<ActionResult> PutAsync(UsuariosSectorProd usuarioSectorProd)
        {
            try
            {
                // Verificar si ya existe una asignación para este usuario y sector (excluyendo el actual)
                var existingAsignacion = await _context.UsuariosSectorProds
                    .FirstOrDefaultAsync(x => x.IdFlexoft == usuarioSectorProd.IdFlexoft && 
                                            x.CodigoSecProd == usuarioSectorProd.CodigoSecProd &&
                                            x.Id != usuarioSectorProd.Id &&
                                            x.Status >= 0);

                if (existingAsignacion != null)
                {
                    return BadRequest($"Ya existe una asignación para el usuario {usuarioSectorProd.NombreUsuario} y el sector {usuarioSectorProd.Descripcion}");
                }

                _context.Update(usuarioSectorProd);
                await _context.SaveChangesAsync();
                return new OkObjectResult(usuarioSectorProd);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Elimina una asignación usuario-sector productivo (borrado físico)
        /// </summary>
        /// <param name="id">ID de la asignación a eliminar</param>
        /// <returns>NoContent o NotFound</returns>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var usuarioSectorProd = await _context.UsuariosSectorProds.FirstOrDefaultAsync(x => x.Id == id);
            if (usuarioSectorProd == null)
            {
                return NotFound();
            }
            
            _context.Remove(usuarioSectorProd);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca una asignación usuario-sector productivo como inactiva (borrado lógico)
        /// </summary>
        /// <param name="id">ID de la asignación</param>
        /// <param name="status">Nuevo status (-1 para inactivo, 0 para activo)</param>
        /// <returns>NoContent o NotFound</returns>
        [HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        {
            var usuarioSectorProd = await _context.UsuariosSectorProds.FirstOrDefaultAsync(x => x.Id == id);
            if (usuarioSectorProd == null)
            {
                return NotFound();
            }

            // Cambiar el campo status 
            usuarioSectorProd.Status = status;

            // Guardar los cambios en la base de datos
            await _context.SaveChangesAsync();

            return NoContent();
        }
        [HttpGet("codigossector/{idUsuario}")]
        public async Task<ActionResult<List<int>>> ObtenerCodigosSector(int idUsuario)
        {
            var codigos = await _context.UsuariosSectorProds
                .Where(u => u.IdFlexoft == idUsuario)
                .Select(u => u.CodigoSecProd)
                .Distinct()
                .ToListAsync();

            if (codigos == null || !codigos.Any())
            {
                return NotFound(new { mensaje = "No se encontraron sectores asignados para el usuario." });
            }

            return Ok(codigos);
        }

    }



}
