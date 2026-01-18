using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class GalimModController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly FileStorageService _fileStorage;
        private readonly ILogger<GalimModController> _logger;

        public GalimModController(DataContext context, FileStorageService fileStorage, ILogger<GalimModController> logger)
        {
            _context = context;
            _fileStorage = fileStorage;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene todas las imágenes de un modelo específico
        /// </summary>
        [HttpGet("modelo/{codMod}")]
        public async Task<ActionResult<List<GalimMod>>> GetByModelo(string codMod)
        {
            try
            {
                var imagenes = await _context.GalimMods
                    .Where(g => g.Cod_Mod == codMod)
                    .OrderBy(g => g.Norit_Imag)
                    .ToListAsync();

                return Ok(imagenes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener imágenes del modelo {CodMod}", codMod);
                return StatusCode(500, "Error al obtener las imágenes");
            }
        }

        /// <summary>
        /// Sube una nueva imagen para un modelo
        /// </summary>
        [HttpPost("upload")]
        public async Task<ActionResult<GalimMod>> UploadImagen([FromForm] IFormFile file, [FromForm] string codMod, [FromForm] string? descripcion)
        {
            try
            {
                // Logging detallado para diagnóstico
                _logger.LogInformation("=== INICIO UploadImagen ===");
                _logger.LogInformation("File: {FileName}, Size: {FileSize}", file?.FileName ?? "NULL", file?.Length ?? 0);
                _logger.LogInformation("CodMod: '{CodMod}'", codMod ?? "NULL");
                _logger.LogInformation("Descripcion: '{Descripcion}'", descripcion ?? "NULL");

                if (file == null || file.Length == 0)
                {
                    _logger.LogWarning("Archivo nulo o vacío");
                    return BadRequest("No se proporcionó ningún archivo");
                }

                if (string.IsNullOrWhiteSpace(codMod))
                {
                    _logger.LogWarning("CodMod nulo o vacío");
                    return BadRequest("No se proporcionó el código del modelo");
                }

                // Validar extensión
                var extensionesPermitidas = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp" };
                var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
                
                _logger.LogInformation("Extensión del archivo: {Extension}", extension);
                
                if (!extensionesPermitidas.Contains(extension))
                {
                    _logger.LogWarning("Extensión no permitida: {Extension}", extension);
                    return BadRequest("Tipo de archivo no permitido. Solo se permiten imágenes.");
                }

                // Obtener el siguiente número de orden
                _logger.LogInformation("Consultando imágenes existentes para modelo: {CodMod}", codMod);
                
                var imagenesExistentes = await _context.GalimMods
                    .Where(g => g.Cod_Mod == codMod)
                    .ToListAsync();
                
                _logger.LogInformation("Imágenes existentes encontradas: {Count}", imagenesExistentes.Count);
                
                short maxNorit = -1;
                if (imagenesExistentes.Any())
                {
                    maxNorit = imagenesExistentes.Max(g => g.Norit_Imag);
                }
                
                short nuevoNorit = (short)(maxNorit + 1);
                _logger.LogInformation("Nuevo Norit_Imag: {NuevoNorit}", nuevoNorit);

                // Crear nombre único para el archivo
                var nombreArchivo = $"Modelo_{codMod}_Img_{nuevoNorit}_{Guid.NewGuid()}{extension}";
                _logger.LogInformation("Nombre de archivo generado: {NombreArchivo}", nombreArchivo);
                
                // Guardar archivo usando FileStorageService
                var (filePath, fileUrl) = await _fileStorage.SaveFileAsync(file, "ModelosImagenes", nombreArchivo);
                
                _logger.LogInformation("Archivo guardado exitosamente en: {FilePath}", filePath);

                // Crear registro en base de datos
                // Los campos de auditoría (AddRecord, LastUpdate, NumUsuar, CodiEmprNet) 
                // se asignan automáticamente por AuditInterceptor
                var galimMod = new GalimMod
                {
                    CodiEmpr = 0,
                    Cod_Mod = codMod,
                    Norit_Imag = nuevoNorit,
                    Nom_Imag = nombreArchivo,
                    RutaImagen = fileUrl,
                    Descripcion = descripcion
                };

                _logger.LogInformation("Agregando registro a base de datos...");
                _context.GalimMods.Add(galimMod);
                
                _logger.LogInformation("Ejecutando SaveChangesAsync...");
                await _context.SaveChangesAsync();
                
                _logger.LogInformation("Imagen guardada exitosamente. ID: {Id}", galimMod.Id);
                return Ok(galimMod);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "ERROR CRÍTICO al subir imagen para modelo {CodMod}. Mensaje: {Message}, StackTrace: {StackTrace}", 
                    codMod ?? "NULL", ex.Message, ex.StackTrace);
                
                if (ex.InnerException != null)
                {
                    _logger.LogError("InnerException: {InnerMessage}", ex.InnerException.Message);
                }
                
                return StatusCode(500, $"Error al subir la imagen: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza la descripción de una imagen
        /// </summary>
        [HttpPut("{id}")]
        public async Task<ActionResult> UpdateDescripcion(int id, [FromBody] string descripcion)
        {
            try
            {
                var galimMod = await _context.GalimMods.FindAsync(id);
                
                if (galimMod == null)
                {
                    return NotFound("Imagen no encontrada");
                }

                // Actualizar descripción
                // Los campos de auditoría (LastUpdate, NumUsuar, CodiEmprNet) 
                // se actualizan automáticamente por AuditInterceptor
                galimMod.Descripcion = descripcion;
                
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar descripción de imagen {Id}", id);
                return StatusCode(500, "Error al actualizar la descripción");
            }
        }

        /// <summary>
        /// Elimina una imagen
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<ActionResult> Delete(int id)
        {
            try
            {
                var galimMod = await _context.GalimMods.FindAsync(id);
                
                if (galimMod == null)
                {
                    return NotFound("Imagen no encontrada");
                }

                // Eliminar archivo físico
                if (!string.IsNullOrEmpty(galimMod.Nom_Imag))
                {
                    _fileStorage.DeleteFile("ModelosImagenes", galimMod.Nom_Imag);
                    _logger.LogInformation("Archivo físico eliminado: {NomImag}", galimMod.Nom_Imag);
                }

                // Eliminar registro de base de datos
                _context.GalimMods.Remove(galimMod);
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al eliminar imagen {Id}", id);
                return StatusCode(500, "Error al eliminar la imagen");
            }
        }
    }
}
