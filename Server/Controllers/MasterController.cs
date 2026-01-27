using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la gestión de clientes (Deudor12)
    /// </summary>
    [ApiController]
    [Route("/api/master")]
    public class MasterController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<MasterController> _logger;
        private readonly ISyncService _syncService;

        public MasterController(DataContext context, ILogger<MasterController> logger, ISyncService syncService)
        {
            _context = context;
            _logger = logger;
            _syncService = syncService;
        }

        /// <summary>
        /// Obtiene una lista paginada de productos (Master) con filtro 
        /// </summary>
        [HttpGet("search")]
        public async Task<ActionResult<List<Master>>> SearchMaster([FromQuery] string filter, [FromQuery] int pageSize = 100)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filter))
                {
                    return new List<Master>();
                }

                // Validar pageSize para evitar consultas excesivas
                if (pageSize > 500) pageSize = 500;
                if (pageSize < 10) pageSize = 10;

                var connection = _context.Database.GetDbConnection();
                await connection.OpenAsync();

                var sql = $@"
                    SELECT TOP {pageSize}
                        Codigo,
                        Codint,
                        Descripcion,
                        Umedida,
                        Status,
                        Cosuni,
                        Monecos,
                        Trazable
                    FROM Master
                    WHERE (Status >= 0 OR Status IS NULL)
                    AND (Codigo LIKE @filter OR Descripcion LIKE @filter)
                    ORDER BY Descripcion";

                using var command = connection.CreateCommand();
                command.CommandText = sql;
                command.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@filter", $"%{filter}%"));

                var results = new List<Master>();
                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    results.Add(new Master
                    {
                        Codint = (int)reader["Codint"],
                        Codigo = reader["Codigo"].ToString(),
                        Descripcion = reader["Descripcion"].ToString(),
                        Umedida = reader["Umedida"].ToString(),
                        Status = reader["Status"] == DBNull.Value ? null : (short?)reader["Status"],
                        Cosuni = reader["Cosuni"] == DBNull.Value ? null : (decimal?)reader["Cosuni"],
                        Monecos = reader["Monecos"] == DBNull.Value ? null : (short?)reader["Monecos"],
                        Trazable = reader["Trazable"] == DBNull.Value ? null : (short?)reader["Trazable"]
                    });
                }

                return Ok(results);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en la búsqueda de productos");
                return StatusCode(500, $"Error en la búsqueda: {ex.Message}");
            }
        }

        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, short status = 0)
        {
            try
            {
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                // var queryable = _context.masters.AsQueryable();
                var queryable = _context.masters.AsNoTracking().AsQueryable();


                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    queryable = queryable.Where(x => x.Descripcion!.ToLower().Contains(pagination.Filter.ToLower()));
                }

                if (status >= 0)
                {
                    queryable = queryable.Where(x => x.Status >= status || x.Status == null);
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == status);
                }

                if (!await queryable.AnyAsync())
                {
                    return Ok(new List<Master>());
                }

                var result = await queryable
                    .OrderBy(x => x.Descripcion)
                    .Paginate(pagination)
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener la lista de productos");
                return StatusCode(500, $"Error al obtener la lista de productos: {ex.Message}");
            }
        }

        /// <summary>
        /// Devuelve el total de páginas para la paginación de producto, considerando el filtro y el status
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, short status = 0)
        {
            try
            {
                // Validar que RecordsNumber sea mayor que cero
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                var queryable = _context.masters.AsQueryable();

                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    queryable = queryable.Where(x => x.Descripcion!.ToLower().Contains(pagination.Filter.ToLower()));
                }

                if (status >= 0)
                {
                    queryable = queryable.Where(x => x.Status >= status || x.Status == null);
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == status);
                }

                // Obtener el conteo total de registros que cumplen con el filtro
                double count = await queryable.CountAsync();

                // Si no hay registros, devolver 1 página (para mostrar el mensaje "No hay registros")
                if (count == 0)
                {
                    return Ok(1);
                }

                // Calcular el número total de páginas
                double totalPages = Math.Ceiling(count / pagination.RecordsNumber);

                // Asegurar que siempre haya al menos una página
                totalPages = Math.Max(1, totalPages);

                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                // Log del error para debugging
                _logger.LogError(ex, "Error al calcular el total de páginas");
                return StatusCode(500, $"Error al calcular el total de páginas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene un producto individual por Id
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var producto = await _context.masters.FirstOrDefaultAsync(x => x.Id == id);
            if (producto == null)
            {
                return NotFound();
            }
            return Ok(producto);
        }
        /// <summary>
        /// Obtiene un producto individual por Codint
        /// </summary>
        [HttpGet("xCodint{codint:int}")]
        public async Task<IActionResult> GetxCodintAsync(int codint)
        {
            var producto = await _context.masters.FirstOrDefaultAsync(x => x.Codint == codint);
            if (producto == null)
            {
                return NotFound();
            }
            return Ok(producto);
        }

        /// <summary>
        /// Crea un nuevo producto
        /// </summary>
        [HttpPost]
        public async Task<IActionResult> PostAsync(Master master)
        {
            // Validar que el código no exista
            if (await _context.masters.AnyAsync(x => x.Codigo == master.Codigo))
            {
                return BadRequest($"Ya existe un producto con el código {master.Codigo}");
            }

            // Iniciar una transacción para asegurar la atomicidad de la operación
            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    // Si codint llega en 0, asignar el siguiente valor disponible
                    if (master.Codint == 0)
                    {
                        // Buscar el valor máximo actual de codint y sumarle 1
                        int maxCodint = await _context.masters.MaxAsync(x => (int?)x.Codint) ?? 0;
                        master.Codint = maxCodint + 1;
                    }

                    _context.Add(master);
                    await _context.SaveChangesAsync();

                    // Confirmar la transacción si todo salió bien
                    await transaction.CommitAsync();

                    // Generar archivo de sincronización si está habilitado
                    try
                    {
                        await _syncService.GenerarArchivoSyncMasterAsync(master.Codint);
                    }
                    catch (Exception exSync)
                    {
                        _logger.LogWarning(exSync, "Error al generar archivo de sincronización (no afecta la operación principal)");
                    }

                    // Devuelve una respuesta HTTP 201 (Created) con la URL y el objeto creado
                    return CreatedAtAction(nameof(GetAsync), new { id = master.Id }, master);
                }
                catch (Exception ex)
                {
                    // Si ocurre algún error, deshacer la transacción
                    await transaction.RollbackAsync();

                    // Devolver un error 500 con el mensaje de la excepción
                    return StatusCode(500, $"Ocurrió un error al crear el producto: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Actualiza un producto existente
        /// </summary>
        [HttpPut]
        public async Task<IActionResult> PutAsync(Master master)
        {
            // Iniciar una transacción para asegurar la atomicidad de la operación
            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    // Verificar si el producto existe
                    var productoExistente = await _context.masters.AnyAsync(x => x.Id == master.Id);
                    if (!productoExistente)
                    {
                        return NotFound($"No se encontró el producto con ID {master.Id}");
                    }

                    _context.Update(master);
                    await _context.SaveChangesAsync();

                    // Confirmar la transacción si todo salió bien
                    await transaction.CommitAsync();

                    // Generar archivo de sincronización si está habilitado
                    try
                    {
                        await _syncService.GenerarArchivoSyncMasterAsync(master.Codint);
                    }
                    catch (Exception exSync)
                    {
                        _logger.LogWarning(exSync, "Error al generar archivo de sincronización (no afecta la operación principal)");
                    }

                    return Ok(master);
                }
                catch (Exception ex)
                {
                    // Si ocurre algún error, deshacer la transacción
                    await transaction.RollbackAsync();

                    // Devolver un error 500 con el mensaje de la excepción
                    return StatusCode(500, $"Ocurrió un error al actualizar el producto: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Cambia el estado de un producto
        /// </summary>
        [HttpPut("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> CambiaStatusAsync(int id, short status)
        {
            // Iniciar una transacción para asegurar la atomicidad de la operación
            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    var master = await _context.masters.FirstOrDefaultAsync(x => x.Id == id);
                    if (master == null)
                    {
                        return NotFound($"No se encontró el producto con ID {id}");
                    }

                    master.Status = status;
                    _context.Update(master);
                    await _context.SaveChangesAsync();

                    // Confirmar la transacción si todo salió bien
                    await transaction.CommitAsync();
                    return NoContent();
                }
                catch (Exception ex)
                {
                    // Si ocurre algún error, deshacer la transacción
                    await transaction.RollbackAsync();

                    // Devolver un error 500 con el mensaje de la excepción
                    return StatusCode(500, $"Ocurrió un error al cambiar el estado del producto: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Actualiza el campo Especificacion de un producto
        /// </summary>
        /// <param name="id">ID del producto</param>
        /// <param name="especificacion">Texto de especificación</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPut("{id}/especificacion")]
        public async Task<IActionResult> UpdateEspecificacionAsync(int id, [FromBody] string especificacion)
        {
            try
            {
                var master = await _context.masters.FirstOrDefaultAsync(m => m.Id == id);
                if (master == null)
                {
                    return NotFound($"No se encontró el producto con ID {id}");
                }

                master.Especificacion = especificacion;
                await _context.SaveChangesAsync();

                return Ok(new { Message = "Especificación actualizada correctamente" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al actualizar la especificación del producto {id}");
                return StatusCode(500, $"Error al actualizar la especificación: {ex.Message}");
            }
        }

        /// <summary>
        /// Busca productos/repuestos en la tabla MASTER por código y descripción
        /// Retorna RepuestoDTO para compatibilidad con el sistema de recepciones
        /// </summary>
        /// <param name="busqueda">Término de búsqueda</param>
        /// <returns>Lista de productos que coinciden con los criterios de búsqueda</returns>
        [HttpGet("buscar")]
        public async Task<ActionResult<List<RepuestoDTO>>> Buscar([FromQuery] string? busqueda)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(busqueda))
                {
                    return BadRequest("El término de búsqueda no puede estar vacío");
                }

                // Buscar en la tabla MASTER por código o descripción
                var productos = await _context.masters
                    .Where(m => (m.Status >= 0 || m.Status == null) &&
                               (m.Codigo!.Contains(busqueda) || 
                                m.Descripcion!.Contains(busqueda)))
                    .OrderBy(m => m.Descripcion)
                    .Take(25)
                    .Select(m => new RepuestoDTO
                    {
                        Codint = m.Codint,
                        CodigoExterno = m.Codigo ?? "",
                        Descritem = m.Descripcion ?? "",
                        Uni_Com = m.Umedida ?? "",
                        UnidadMedidaDescripcion = m.Umedida ?? ""
                    })
                    .ToListAsync();

                _logger.LogInformation($"Productos encontrados en MASTER: {productos.Count}");
                return Ok(productos);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en búsqueda de productos: {ex.Message}");
                _logger.LogError(ex, "Error al buscar productos en MASTER");
                return StatusCode(500, $"Error interno al realizar la búsqueda: {ex.Message}");
            }
        }
        [HttpGet("GetMaster1Dto")]
        public async Task<IActionResult> GetMaster1Dto([FromQuery] short status = 0, [FromQuery] string? filter = null)
        {
            try
            {
                var queryable = _context.masters.AsNoTracking().AsQueryable();

                if (status >= 0)
                {
                    queryable = queryable.Where(x => x.Status >= status || x.Status == null);
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == status);
                }

                // 🔎 Agregamos el filtro textual si fue enviado
                if (!string.IsNullOrWhiteSpace(filter))
                {
                    string filtro = filter.Trim().ToLower();
                    queryable = queryable.Where(x =>
                        x.Descripcion!.ToLower().Contains(filtro) ||
                        x.Codigo!.ToLower().Contains(filtro)
                    // podés agregar más campos si querés
                    );
                }

                if (!await queryable.AnyAsync())
                {
                    return Ok(new List<Master1DTO>());
                }

                var result = await queryable
                    .OrderBy(x => x.Descripcion)
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener la lista de productos");
                return StatusCode(500, $"Error al obtener la lista de productos: {ex.Message}");
            }
        }

        [HttpGet("cantidad-registros")]
        public async Task<ActionResult<int>> GetCantidadRegistros([FromQuery] string? filter = "", short status = 0)
        {
            try
            {
                var queryable = _context.masters.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    queryable = queryable.Where(x => x.Descripcion!.ToLower().Contains(filter.ToLower()));
                }

                if (status >= 0)
                {
                    queryable = queryable.Where(x => x.Status >= status || x.Status == null);
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == status);
                }

                int count = await queryable.CountAsync();
                return Ok(count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener cantidad de registros");
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }
        //listqado de pantalla
        [HttpGet("GetMasterListado")]
        public async Task<IActionResult> GetMasterListado([FromQuery] short status = 0, [FromQuery] string? filter = null)
        {
            try
            {
                // Carga los tipos de material en memoria
                var tatimats = await _context.Tatimat
                    .AsNoTracking()
                    .ToListAsync();

                // Carga los masters en memoria aplicando los filtros
                var mastersFiltrados = await _context.masters
                    .AsNoTracking()
                    .Where(x =>
                        (status >= 0 ? x.Status >= status || x.Status == null : x.Status == status) &&
                        (string.IsNullOrWhiteSpace(filter) ||
                         (x.Descripcion ?? "").ToLower().Contains(filter.Trim().ToLower()) ||
                         (x.Codigo ?? "").ToLower().Contains(filter.Trim().ToLower()))
                    )
                    .OrderBy(x => x.Descripcion)
                    .ToListAsync();

                // Proyección en memoria usando los datos ya cargados
                var resultado = mastersFiltrados.Select(m => new MasterListadoIndexDTO
                {
                    Id = m.Id.ToString(),
                    Codint = m.Codint.ToString(),
                    Codigo = m.Codigo,
                    Descripcion = m.Descripcion,
                    Umedida = m.Umedida,
                    Timate = m.Timate.ToString(),
                    TipoMaterialDescripcion = m.Timate != null
                        ? tatimats.FirstOrDefault(t => t.CodigoTatimat == m.Timate)?.Descripcion ?? "-"
                        : "-"
                }).OrderBy(x => x.Descripcion).ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener MasterListadoDTO");
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }
    }
}



//CORRECCIONES DE LA TABLA MASTER
//SE CORRIO EL SCRIPT PARA CREAR LA TABLA EN ARCON2 CON EL ESQUEMA Y LOS DATOS Y LA GENERO DE 0 CORRECTAMENTE
//LUEGO SE AGREGO UN CAMPO IDENTITY
//ALTER TABLE MASTER ADD ID INT IDENTITY(1,1);

//ESTO PARA CAMBIAR LOS DATOS DE FLOAT A DECIMAL VERIFICAR SI EL CAMBIO NO AFECTA AL SISTEMA FLEXOFT
//DECLARE @sql NVARCHAR(MAX);

//SELECT @sql = STRING_AGG('ALTER TABLE MASTER ALTER COLUMN ' + COLUMN_NAME + ' DECIMAL(18,2);', ' ')
//FROM INFORMATION_SCHEMA.COLUMNS
//WHERE TABLE_NAME = 'MASTER' AND DATA_TYPE = 'float'; --Cambia 'float' si tu tipo es 'double'

//EXEC sp_executesql @sql;
//   ***** OJO VER LOS CAMPOS QUE USAN 4 DECIMALES ***** SI NO CAMBIARLOS MANUALMENTE, TAMBIEN CORREGIR LA ENTIDAD
// EN *** ARCON **** SE USO DE ESTA MANERA PARA PASAR A DECIMAL LOS FLOAT
//DECLARE @sql NVARCHAR(MAX) = '';

//SELECT @sql = @sql + 'ALTER TABLE MASTER ALTER COLUMN ' + COLUMN_NAME + ' DECIMAL(18,2); ' 
//FROM INFORMATION_SCHEMA.COLUMNS 
//WHERE TABLE_NAME = 'MASTER' AND DATA_TYPE = 'float';

//EXEC sp_executesql @sql;
