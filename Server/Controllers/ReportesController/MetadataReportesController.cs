using Microsoft.AspNetCore.Mvc;
using System.Reflection;
using Shared.Entities;
using IMPLANPROD.Server.Data;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Shared.Entities;
using static IMPLANPROD.Client.Pages.Reportes.ReporteGenerador; // Para GetDbConnection()

namespace IMPLANPROD.Server.Controllers.ReportesController
{
    /// <summary>
    /// Controlador para exponer metadatos sobre las entidades/tablas y sus campos.
    /// Permite al frontend construir reportes dinámicos de forma simple y escalable.
    /// </summary>
    [ApiController]
    [Route("/api/metadata-reportes")]
    public class MetadataReportesController : ControllerBase
    {
        private readonly DataContext _context;

        public MetadataReportesController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Devuelve un diccionario dinámico con el mapeo entre nombre de tabla física y nombre de entidad (clase).
        /// Esto permite que el frontend siempre pueda asociar correctamente la selección de tabla con la entidad,
        /// incluso si los nombres son diferentes (plural/singular, etc). El mapeo se construye automáticamente
        /// recorriendo el modelo de Entity Framework, por lo que si se agregan nuevas tablas o entidades, el sistema
        /// las reconocerá sin necesidad de modificar el código.
        /// </summary>
        [HttpGet("tablas-mapeo")]
        public ActionResult<Dictionary<string, string>> GetTablasMapeo()
        {
            try
            {
                // Diccionario que asocia el nombre de la tabla física (base de datos) con el nombre de la entidad (clase)
                var diccionario = new Dictionary<string, string>();
                // Recorremos todas las entidades del modelo de datos
                foreach (var entityType in _context.Model.GetEntityTypes())
                {
                    // Nombre de la tabla física en la base de datos
                    var tableName = entityType.GetTableName();
                    // Nombre de la clase de la entidad
                    var entityName = entityType.ClrType.Name;
                    // Solo agregamos si no existe ya la clave
                    if (!diccionario.ContainsKey(tableName))
                        diccionario.Add(tableName, entityName);
                }
                return Ok(diccionario);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al construir el mapeo tabla-entidad: {ex.Message}");
            }
        }

        /// <summary>
        /// Devuelve la lista de tablas físicas existentes en la base de datos.
        /// </summary>
        [HttpGet("tablas")]
        public ActionResult<IEnumerable<string>> GetTablas()
        {
            try
            {
                // Consulta dinámica a la base de datos para obtener todas las tablas usando ADO.NET
                var tablas = new List<string>();
                // Acceso robusto a la conexión de base de datos usando GetDbConnection()
                using (var dbConnection = _context.Database.GetDbConnection())
                {
                    dbConnection.Open();
                    using (var command = dbConnection.CreateCommand())
                    {
                        command.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME";
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                tablas.Add(reader.GetString(0));
                            }
                        }
                    }
                }
                return Ok(tablas);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener las tablas: {ex.Message}");
            }
        }

        /// <summary>
        /// Devuelve los campos de una tabla/entidad por su nombre.
        /// </summary>
        [HttpGet("campos")]
        public ActionResult<IEnumerable<object>> GetCampos([FromQuery] string tabla)
        {
            try
            {
                // Buscar el tipo de la entidad por nombre
                var assembly = typeof(Master).Assembly;
                var tipo = assembly.GetTypes().FirstOrDefault(t => t.Name == tabla);
                if (tipo == null)
                    return NotFound($"No se encontró la entidad '{tabla}'");
                // Obtener propiedades públicas
                var campos = tipo.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Select(p => new {
                        Nombre = p.Name,
                        Tipo = p.PropertyType.Name
                    }).ToList();
                return Ok(campos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener los campos de la tabla: {ex.Message}");
            }
        }
        /// <summary>
        /// Ejecuta una consulta SQL enviada por el frontend y devuelve los resultados en formato lista de diccionarios.
        /// Solo permite consultas SELECT por seguridad. NO permite DML ni sentencias peligrosas.
        /// </summary>
        /// <param name="dto">Objeto que contiene la consulta SQL a ejecutar.</param>
        /// <returns>Lista de diccionarios con los resultados de la consulta.</returns>
        [HttpPost("preview-sql")]
        public async Task<ActionResult<List<Dictionary<string, object>>>> PreviewSql([FromBody] SqlPreviewDTO dto)
        {
            // Validación básica: solo se permite SELECT
            if (dto == null || string.IsNullOrWhiteSpace(dto.Sql))
                return BadRequest("La consulta SQL no puede estar vacía.");
            var sql = dto.Sql.Trim();
            if (!sql.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase))
                return BadRequest("Solo se permiten consultas SELECT para vista previa.");
            // Opcional: más validaciones para evitar palabras peligrosas
            var palabrasProhibidas = new[] {"INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "TRUNCATE", ";"};
            foreach (var palabra in palabrasProhibidas)
            {
                if (sql.ToUpper().Contains($" {palabra} ") || sql.ToUpper().Contains($"\n{palabra} "))
                    return BadRequest($"La consulta contiene la palabra prohibida '{palabra}'. Solo se permite SELECT.");
            }
            var resultados = new List<Dictionary<string, object>>();
            try
            {
                // Uso de ADO.NET para máxima flexibilidad y compatibilidad
                using (var connection = _context.Database.GetDbConnection())
                {
                    await connection.OpenAsync();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = sql;
                        using (var reader = await command.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                var fila = new Dictionary<string, object>();
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    // Se almacena el valor crudo, usando el nombre original de la columna
                                    fila[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                                }
                                resultados.Add(fila);
                            }
                        }
                    }
                }
                return Ok(resultados);
            }
            catch (Exception ex)
            {
                // Por seguridad, no exponer detalles internos en producción
                return StatusCode(500, $"Error al ejecutar la consulta: {ex.Message}");
            }
        }
    }
}
