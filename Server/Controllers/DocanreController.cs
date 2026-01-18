using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocanreController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly FileStorageService _fileStorage;
        private readonly ILogger<DocanreController> _logger;

        private static string GetFullExceptionMessage(Exception ex)
        {
            var messages = new List<string>();
            Exception? current = ex;
            while (current != null)
            {
                if (!string.IsNullOrWhiteSpace(current.Message))
                {
                    messages.Add(current.Message);
                }

                current = current.InnerException;
            }

            return string.Join(" | ", messages);
        }

        public DocanreController(DataContext context, FileStorageService fileStorage, ILogger<DocanreController> logger)
        {
            _context = context;
            _fileStorage = fileStorage;
            _logger = logger;
        }

        [HttpGet("reparacion/{numRepa}")]
        public async Task<ActionResult<List<Docanre>>> GetByReparacion(int numRepa)
        {
            try
            {
                var documentos = await _context.Docanres
                    .Where(d => d.NumRepa == numRepa)
                    .OrderByDescending(d => d.AddRecord)
                    .ThenByDescending(d => d.Id)
                    .ToListAsync();

                return Ok(documentos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener documentos de la reparación {NumRepa}", numRepa);
                return StatusCode(500, $"Error al obtener los documentos. Detalle: {GetFullExceptionMessage(ex)}");
            }
        }

        [HttpPost("upload")]
        public async Task<ActionResult<Docanre>> UploadDocumento(
            [FromForm] IFormFile file,
            [FromForm] int numRepa)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest("No se proporcionó ningún archivo");
                }

                var extension = Path.GetExtension(file.FileName).ToLowerInvariant();

                // Lista básica de extensiones permitidas (expandible)
                var extensionesPermitidas = new[]
                {
                    ".pdf", ".doc", ".docx", ".txt", ".rtf",
                    ".xls", ".xlsx", ".csv",
                    ".png", ".jpg", ".jpeg"
                };

                if (string.IsNullOrWhiteSpace(extension) || !extensionesPermitidas.Contains(extension))
                {
                    return BadRequest("Tipo de archivo no permitido.");
                }

                // Guardar archivo usando FileStorageService
                var (filePath, fileUrl) = await _fileStorage.SaveFileAsync(file, "RepalDoc", $"Repa_{numRepa}_Doc_{Guid.NewGuid()}{extension}");

                // Los campos de auditoría se asignan automáticamente por AuditInterceptor
                var doc = new Docanre
                {
                    CodiEmpr = 0,
                    NumRepa = numRepa,
                    NomAsignado = Path.GetFileName(filePath),
                    NomOriginal = file.FileName,
                    RutaDocumento = fileUrl,
                    ContentType = file.ContentType,
                    TamanoBytes = file.Length,
                    Extension = extension
                };

                _context.Docanres.Add(doc);
                await _context.SaveChangesAsync();

                return Ok(doc);
            }
            catch (UnauthorizedAccessException ex)
            {
                _logger.LogError(ex, "Acceso denegado al intentar escribir archivo para reparación {NumRepa}", numRepa);
                return StatusCode(500, $"No se pudo subir el documento por falta de permisos de escritura en el servidor. Detalle: {GetFullExceptionMessage(ex)}");
            }
            catch (IOException ex)
            {
                _logger.LogError(ex, "Error de E/S al intentar escribir archivo para reparación {NumRepa}", numRepa);
                return StatusCode(500, $"No se pudo subir el documento por un error de acceso a disco (E/S). Detalle: {GetFullExceptionMessage(ex)}");
            }
            catch (DbUpdateException ex)
            {
                _logger.LogError(ex, "Error de base de datos al registrar documento para reparación {NumRepa}", numRepa);
                return StatusCode(500, $"No se pudo registrar el documento en la base de datos. Verifique que exista la tabla DOCANRE y reintente. Detalle: {GetFullExceptionMessage(ex)}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al subir documento para reparación {NumRepa}", numRepa);
                return StatusCode(500, $"No se pudo subir el documento. Detalle: {GetFullExceptionMessage(ex)}");
            }
        }

        [HttpGet("download/{id}")]
        public async Task<IActionResult> Download(int id)
        {
            var doc = await _context.Docanres.FindAsync(id);
            if (doc == null)
            {
                return NotFound("Documento no encontrado");
            }

            if (string.IsNullOrWhiteSpace(doc.NomAsignado))
            {
                return NotFound("Documento sin nombre asignado");
            }

            var rutaCompleta = _fileStorage.GetFilePath("RepalDoc", doc.NomAsignado);
            if (!System.IO.File.Exists(rutaCompleta))
            {
                return NotFound("Archivo físico no encontrado");
            }

            var downloadName = string.IsNullOrWhiteSpace(doc.NomOriginal)
                ? doc.NomAsignado
                : doc.NomOriginal;

            return File(System.IO.File.OpenRead(rutaCompleta), doc.ContentType ?? "application/octet-stream", downloadName);
        }

        [HttpGet("view/{id}")]
        public async Task<IActionResult> View(int id)
        {
            var doc = await _context.Docanres.FindAsync(id);
            if (doc == null)
            {
                return NotFound("Documento no encontrado");
            }

            if (string.IsNullOrWhiteSpace(doc.NomAsignado))
            {
                return NotFound("Documento sin nombre asignado");
            }

            if (!_fileStorage.FileExists("RepalDoc", doc.NomAsignado))
            {
                return NotFound("Archivo físico no encontrado");
            }

            var rutaCompleta = _fileStorage.GetFilePath("RepalDoc", doc.NomAsignado);
            var contentType = string.IsNullOrWhiteSpace(doc.ContentType)
                ? "application/octet-stream"
                : doc.ContentType;

            var fileName = string.IsNullOrWhiteSpace(doc.NomOriginal)
                ? doc.NomAsignado
                : doc.NomOriginal;

            Response.Headers["Content-Disposition"] = $"inline; filename=\"{fileName}\"";
            return PhysicalFile(rutaCompleta, contentType);
        }

        [HttpDelete("{id}")]
        public async Task<ActionResult> Delete(int id)
        {
            try
            {
                var doc = await _context.Docanres.FindAsync(id);
                if (doc == null)
                {
                    return NotFound("Documento no encontrado");
                }

                if (!string.IsNullOrWhiteSpace(doc.NomAsignado))
                {
                    _fileStorage.DeleteFile("RepalDoc", doc.NomAsignado);
                }

                _context.Docanres.Remove(doc);
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al eliminar documento {Id}", id);
                return StatusCode(500, $"Error al eliminar el documento. Detalle: {GetFullExceptionMessage(ex)}");
            }
        }
    }
}
