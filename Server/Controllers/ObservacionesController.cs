using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using System.Reflection;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador genérico para gestionar observaciones de diferentes tablas
    /// Permite obtener y guardar observaciones en tablas con estructura similar
    /// Usa Entity Framework para aprovechar el AuditInterceptor automático
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class ObservacionesController : ControllerBase
    {
        private readonly DataContext _context;

        public ObservacionesController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene las observaciones de BADAMODE_NOTA filtradas por código de modelo
        /// </summary>
        /// <param name="tabla">Nombre de la tabla (solo BADAMODE_NOTA soportado)</param>
        /// <param name="campo">Nombre del campo identificador (Cod_Mod)</param>
        /// <param name="valor">Valor del código de modelo</param>
        /// <param name="esTexto">Indica si el campo es texto (true) o numérico (false)</param>
        /// <returns>Lista de observaciones ordenadas por fecha de creación descendente</returns>
        [HttpGet]
        public async Task<ActionResult<List<ObservacionDTO>>> GetObservaciones(
            [FromQuery] string tabla,
            [FromQuery] string campo,
            [FromQuery] string valor,
            [FromQuery] bool esTexto = true)
        {
            try
            {
                // Validar parámetros
                if (string.IsNullOrWhiteSpace(tabla) || 
                    string.IsNullOrWhiteSpace(campo) || 
                    string.IsNullOrWhiteSpace(valor))
                {
                    return BadRequest("Los parámetros tabla, campo y valor son obligatorios");
                }

                // Por ahora solo soportamos BADAMODE_NOTA usando Entity Framework
                if (tabla.ToUpper() != "BADAMODE_NOTA")
                {
                    return BadRequest($"Tabla '{tabla}' no soportada. Use BADAMODE_NOTA");
                }

                // Usar Entity Framework para obtener las observaciones
                var observaciones = await _context.BadamodeNotas
                    .Where(n => n.Cod_Mod == valor)
                    .OrderByDescending(n => n.AddRecord)
                    .Select(n => new ObservacionDTO
                    {
                        Id = n.Id,
                        AddRecord = n.AddRecord,
                        Observaciones = n.Observaciones,
                        NumUsuar = n.NumUsuar
                    })
                    .ToListAsync();

                return Ok(observaciones);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener observaciones: {ex.Message}");
                return StatusCode(500, $"Error al obtener observaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Guarda una nueva observación en BADAMODE_NOTA
        /// Los campos de auditoría (AddRecord, LastUpdate, NumUsuar, CodiEmprNet) 
        /// se llenan automáticamente por el AuditInterceptor
        /// </summary>
        /// <param name="dto">Datos de la observación a guardar</param>
        /// <returns>La observación creada</returns>
        [HttpPost]
        public async Task<ActionResult<ObservacionDTO>> PostObservacion([FromBody] GuardarObservacionDTO dto)
        {
            try
            {
                // Validar parámetros
                if (string.IsNullOrWhiteSpace(dto.NombreTabla) || 
                    string.IsNullOrWhiteSpace(dto.CampoIdentificador) || 
                    string.IsNullOrWhiteSpace(dto.ValorIdentificador) ||
                    string.IsNullOrWhiteSpace(dto.Observaciones))
                {
                    return BadRequest("Todos los campos son obligatorios");
                }

                // Por ahora solo soportamos BADAMODE_NOTA usando Entity Framework
                if (dto.NombreTabla.ToUpper() != "BADAMODE_NOTA")
                {
                    return BadRequest($"Tabla '{dto.NombreTabla}' no soportada. Use BADAMODE_NOTA");
                }

                Console.WriteLine($"[ObservacionesController] Guardando observación para {dto.CampoIdentificador}={dto.ValorIdentificador}");

                // Crear nueva entidad BadamodeNota
                // El AuditInterceptor llenará automáticamente: AddRecord, LastUpdate, NumUsuar, CodiEmprNet
                var nuevaNota = new BadamodeNota
                {
                    Cod_Mod = dto.ValorIdentificador,
                    Observaciones = dto.Observaciones
                    // NO asignar AddRecord, LastUpdate, NumUsuar, CodiEmprNet
                    // El AuditInterceptor los llenará automáticamente
                };

                // Agregar a DbSet
                _context.BadamodeNotas.Add(nuevaNota);

                // Guardar cambios - aquí se ejecuta el AuditInterceptor
                await _context.SaveChangesAsync();

                Console.WriteLine($"[ObservacionesController] Observación guardada con ID: {nuevaNota.Id}");
                Console.WriteLine($"[ObservacionesController] AddRecord: {nuevaNota.AddRecord}");
                Console.WriteLine($"[ObservacionesController] LastUpdate: {nuevaNota.LastUpdate}");
                Console.WriteLine($"[ObservacionesController] NumUsuar: {nuevaNota.NumUsuar}");
                Console.WriteLine($"[ObservacionesController] CodiEmprNet: {nuevaNota.CodiEmprNet}");

                // Retornar la observación creada con todos los campos llenados por el interceptor
                var observacionCreada = new ObservacionDTO
                {
                    Id = nuevaNota.Id,
                    AddRecord = nuevaNota.AddRecord,
                    Observaciones = nuevaNota.Observaciones,
                    NumUsuar = nuevaNota.NumUsuar
                };

                return Ok(observacionCreada);
            }
            catch (UnauthorizedAccessException ex)
            {
                // Error del AuditInterceptor cuando no hay sesión válida
                Console.WriteLine($"[ObservacionesController] Error de auditoría: {ex.Message}");
                return StatusCode(401, ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ObservacionesController] Error al guardar observación: {ex.Message}");
                Console.WriteLine($"[ObservacionesController] StackTrace: {ex.StackTrace}");
                return StatusCode(500, $"Error al guardar observación: {ex.Message}");
            }
        }
    }
}
