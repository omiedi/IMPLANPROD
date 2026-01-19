using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la gestión de Modelos (BADAMODE)
    /// Proporciona endpoints CRUD para la administración de moldes
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class BadamodeController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<BadamodeController> _logger;

        public BadamodeController(DataContext context, ILogger<BadamodeController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene una lista paginada de modelos activos
        /// </summary>
        /// <param name="page">Número de página</param>
        /// <param name="filter">Filtro de búsqueda por código o descripción</param>
        /// <param name="recordsNumber">Número de registros por página (opcional, para obtener todos usar 999999)</param>
        /// <returns>Lista de modelos</returns>
        [HttpGet]
        public async Task<ActionResult<List<Badamode>>> GetAsync([FromQuery] int page = 1, [FromQuery] string? filter = null, [FromQuery] int? recordsNumber = null)
        {
            try
            {
                var recordsPerPage = recordsNumber ?? 10; // Usar recordsNumber si se proporciona, sino 10 por defecto
                var query = _context.Badamodes
                    .Where(m => m.Status != -1) // Excluir modelos dados de baja
                    .AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    query = query.Where(m => 
                        (m.Cod_Mod != null && m.Cod_Mod.Contains(filter)) ||
                        (m.Descripcion != null && m.Descripcion.Contains(filter)));
                }

                var modelos = await query
                    .OrderBy(m => m.Cod_Mod)
                    .Skip((page - 1) * recordsPerPage)
                    .Take(recordsPerPage)
                    .ToListAsync();

                _logger.LogInformation("Obtenidos {Count} modelos (página {Page}, registros por página: {RecordsPerPage})", 
                    modelos.Count, page, recordsPerPage);

                return Ok(modelos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener modelos");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        /// <param name="filter">Filtro de búsqueda</param>
        /// <returns>Número total de páginas</returns>
        [HttpGet("totalPages")]
        public async Task<ActionResult<int>> GetPagesAsync([FromQuery] string? filter = null)
        {
            try
            {
                var recordsPerPage = 10;
                var query = _context.Badamodes
                    .Where(m => m.Status != -1)
                    .AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    query = query.Where(m => 
                        (m.Cod_Mod != null && m.Cod_Mod.Contains(filter)) ||
                        (m.Descripcion != null && m.Descripcion.Contains(filter)));
                }

                var totalRecords = await query.CountAsync();
                var totalPages = (int)Math.Ceiling((double)totalRecords / recordsPerPage);

                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener total de páginas");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene un modelo por su ID
        /// </summary>
        /// <param name="id">ID del modelo</param>
        /// <returns>Modelo encontrado</returns>
        [HttpGet("{id:int}")]
        public async Task<ActionResult<Badamode>> GetByIdAsync(int id)
        {
            try
            {
                var modelo = await _context.Badamodes
                    .FirstOrDefaultAsync(m => m.Id == id);

                if (modelo == null)
                {
                    return NotFound($"Modelo con ID {id} no encontrado");
                }

                return Ok(modelo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener modelo por ID {Id}", id);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene un modelo por su código
        /// </summary>
        /// <param name="codigo">Código del modelo</param>
        /// <returns>Modelo encontrado</returns>
        [HttpGet("codigo/{codigo}")]
        public async Task<ActionResult<Badamode>> GetByCodigoAsync(string codigo)
        {
            try
            {
                var modelo = await _context.Badamodes
                    .FirstOrDefaultAsync(m => m.Cod_Mod == codigo);

                if (modelo == null)
                {
                    return NotFound($"Modelo con código {codigo} no encontrado");
                }

                return Ok(modelo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener modelo por código {Codigo}", codigo);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Crea un nuevo modelo
        /// </summary>
        /// <param name="modelo">Datos del modelo a crear</param>
        /// <returns>Modelo creado</returns>
        [HttpPost]
        public async Task<ActionResult<Badamode>> PostAsync([FromBody] Badamode modelo)
        {
            try
            {
                // Validar que no exista un modelo con el mismo código
                var existeCodigo = await _context.Badamodes
                    .AnyAsync(m => m.Cod_Mod == modelo.Cod_Mod && m.Status != -1);

                if (existeCodigo)
                {
                    return BadRequest($"Ya existe un modelo con el código {modelo.Cod_Mod}");
                }

                // Establecer valores por defecto
                modelo.AddRecord = DateTime.Now;
                modelo.LastUpdate = DateTime.Now;
                modelo.Status = modelo.Status ?? 1; // Activo por defecto

                _context.Badamodes.Add(modelo);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Modelo creado: {Codigo}", modelo.Cod_Mod);

                return CreatedAtAction(nameof(GetByIdAsync), new { id = modelo.Id }, modelo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear modelo");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Actualiza un modelo existente
        /// </summary>
        /// <param name="id">ID del modelo</param>
        /// <param name="modelo">Datos actualizados del modelo</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPut("{id:int}")]
        public async Task<ActionResult> PutAsync(int id, [FromBody] Badamode modelo)
        {
            try
            {
                if (id != modelo.Id)
                {
                    return BadRequest("El ID del modelo no coincide");
                }

                var modeloExistente = await _context.Badamodes
                    .FirstOrDefaultAsync(m => m.Id == id);

                if (modeloExistente == null)
                {
                    return NotFound($"Modelo con ID {id} no encontrado");
                }

                // Validar que no exista otro modelo con el mismo código
                var existeCodigo = await _context.Badamodes
                    .AnyAsync(m => m.Cod_Mod == modelo.Cod_Mod && m.Id != id && m.Status != -1);

                if (existeCodigo)
                {
                    return BadRequest($"Ya existe otro modelo con el código {modelo.Cod_Mod}");
                }

                // Actualizar campos
                modeloExistente.Cod_Mod = modelo.Cod_Mod;
                modeloExistente.CodiEmpr = modelo.CodiEmpr;
                modeloExistente.Descripcion = modelo.Descripcion;
                modeloExistente.Ubic = modelo.Ubic;
                modeloExistente.CantPiezas = modelo.CantPiezas;
                modeloExistente.Frecuen = modelo.Frecuen;
                modeloExistente.FechUlt_Con = modelo.FechUlt_Con;
                modeloExistente.FechPro_Con = modelo.FechPro_Con;
                modeloExistente.Material = modelo.Material;
                modeloExistente.Desc_Mat = modelo.Desc_Mat;
                modeloExistente.Cod_Inte = modelo.Cod_Inte;
                modeloExistente.ObservaMod = modelo.ObservaMod;
                modeloExistente.Foto1 = modelo.Foto1;
                modeloExistente.Foto2 = modelo.Foto2;
                modeloExistente.Foto3 = modelo.Foto3;
                modeloExistente.Foto4 = modelo.Foto4;
                modeloExistente.Foto_2 = modelo.Foto_2;
                modeloExistente.Foto_3 = modelo.Foto_3;
                modeloExistente.Foto_4 = modelo.Foto_4;
                modeloExistente.Status = modelo.Status;
                modeloExistente.LastUpdate = DateTime.Now;
                modeloExistente.CodiEmprNet = modelo.CodiEmprNet;
                modeloExistente.NumUsuar = modelo.NumUsuar;

                _context.Badamodes.Update(modeloExistente);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Modelo actualizado: {Codigo}", modelo.Cod_Mod);

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar modelo {Id}", id);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Da de baja un modelo (cambia Status a -1)
        /// </summary>
        /// <param name="id">ID del modelo</param>
        /// <returns>Resultado de la operación</returns>
        [HttpDelete("{id:int}")]
        public async Task<ActionResult> DeleteAsync(int id)
        {
            try
            {
                var modelo = await _context.Badamodes
                    .FirstOrDefaultAsync(m => m.Id == id);

                if (modelo == null)
                {
                    return NotFound($"Modelo con ID {id} no encontrado");
                }

                // Baja lógica
                modelo.Status = -1;
                modelo.LastUpdate = DateTime.Now;

                _context.Badamodes.Update(modelo);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Modelo dado de baja: {Codigo}", modelo.Cod_Mod);

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al dar de baja modelo {Id}", id);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene modelos dados de baja (Status = -1)
        /// </summary>
        /// <param name="page">Número de página</param>
        /// <param name="filter">Filtro de búsqueda</param>
        /// <returns>Lista de modelos dados de baja</returns>
        [HttpGet("baja")]
        public async Task<ActionResult<List<Badamode>>> GetBajaAsync([FromQuery] int page = 1, [FromQuery] string? filter = null)
        {
            try
            {
                var recordsPerPage = 10;
                var query = _context.Badamodes
                    .Where(m => m.Status == -1)
                    .AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    query = query.Where(m => 
                        (m.Cod_Mod != null && m.Cod_Mod.Contains(filter)) ||
                        (m.Descripcion != null && m.Descripcion.Contains(filter)));
                }

                var modelos = await query
                    .OrderBy(m => m.Cod_Mod)
                    .Skip((page - 1) * recordsPerPage)
                    .Take(recordsPerPage)
                    .ToListAsync();

                return Ok(modelos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener modelos dados de baja");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para modelos dados de baja
        /// </summary>
        /// <param name="filter">Filtro de búsqueda</param>
        /// <returns>Número total de páginas</returns>
        [HttpGet("baja/totalPages")]
        public async Task<ActionResult<int>> GetBajaPagesAsync([FromQuery] string? filter = null)
        {
            try
            {
                var recordsPerPage = 10;
                var query = _context.Badamodes
                    .Where(m => m.Status == -1)
                    .AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    query = query.Where(m => 
                        (m.Cod_Mod != null && m.Cod_Mod.Contains(filter)) ||
                        (m.Descripcion != null && m.Descripcion.Contains(filter)));
                }

                var totalRecords = await query.CountAsync();
                var totalPages = (int)Math.Ceiling((double)totalRecords / recordsPerPage);

                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener total de páginas de modelos dados de baja");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Reactiva un modelo dado de baja (cambia Status a 1)
        /// </summary>
        /// <param name="id">ID del modelo</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPut("reactivar/{id:int}")]
        public async Task<ActionResult> ReactivarAsync(int id)
        {
            try
            {
                var modelo = await _context.Badamodes
                    .FirstOrDefaultAsync(m => m.Id == id);

                if (modelo == null)
                {
                    return NotFound($"Modelo con ID {id} no encontrado");
                }

                // Reactivar
                modelo.Status = 1;
                modelo.LastUpdate = DateTime.Now;

                _context.Badamodes.Update(modelo);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Modelo reactivado: {Codigo}", modelo.Cod_Mod);

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al reactivar modelo {Id}", id);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene todos los modelos activos sin paginación (para impresión/exportación)
        /// </summary>
        /// <param name="filter">Filtro de búsqueda por código o descripción</param>
        /// <returns>Lista completa de modelos</returns>
        [HttpGet("todos")]
        public async Task<ActionResult<List<Badamode>>> GetTodosAsync([FromQuery] string? filter = null)
        {
            try
            {
                var query = _context.Badamodes
                    .Where(m => m.Status != -1) // Excluir modelos dados de baja
                    .AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    query = query.Where(m => 
                        (m.Cod_Mod != null && m.Cod_Mod.Contains(filter)) ||
                        (m.Descripcion != null && m.Descripcion.Contains(filter)));
                }

                var modelos = await query
                    .OrderBy(m => m.Cod_Mod)
                    .ToListAsync();

                _logger.LogInformation("Obtenidos {Count} modelos para impresión/exportación", modelos.Count);

                return Ok(modelos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener todos los modelos");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene las notas de un modelo específico
        /// </summary>
        /// <param name="codMod">Código del modelo</param>
        /// <returns>Lista de notas del modelo</returns>
        [HttpGet("{codMod}/notas")]
        public async Task<ActionResult<List<BadamodeNota>>> GetNotasAsync(string codMod)
        {
            try
            {
                var notas = await _context.BadamodeNotas
                    .Where(n => n.Cod_Mod == codMod)
                    .OrderByDescending(n => n.AddRecord)
                    .ToListAsync();

                return Ok(notas);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener notas del modelo {CodMod}", codMod);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Agrega una nota a un modelo
        /// </summary>
        /// <param name="nota">Datos de la nota</param>
        /// <returns>Nota creada</returns>
        [HttpPost("notas")]
        public async Task<ActionResult<BadamodeNota>> PostNotaAsync([FromBody] BadamodeNota nota)
        {
            try
            {
                nota.AddRecord = DateTime.Now;
                nota.LastUpdate = DateTime.Now;

                _context.BadamodeNotas.Add(nota);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Nota agregada al modelo {CodMod}", nota.Cod_Mod);

                return CreatedAtAction(nameof(GetNotasAsync), new { codMod = nota.Cod_Mod }, nota);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al agregar nota");
                return StatusCode(500, "Error interno del servidor");
            }
        }
    }
}
