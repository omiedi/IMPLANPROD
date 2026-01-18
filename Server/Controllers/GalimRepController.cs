using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class GalimRepController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly FileStorageService _fileStorage;
        private readonly ILogger<GalimRepController> _logger;

        public GalimRepController(DataContext context, FileStorageService fileStorage, ILogger<GalimRepController> logger)
        {
            _context = context;
            _fileStorage = fileStorage;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene todas las imágenes de una reparación específica
        /// </summary>
        [HttpGet("reparacion/{numRepa}")]
        public async Task<ActionResult<List<GalimRep>>> GetByReparacion(int numRepa)
        {
            try
            {
                var imagenes = await _context.GalimReps
                    .Where(g => g.NumRepa == numRepa)
                    .OrderBy(g => g.Norit_Imag)
                    .ToListAsync();

                return Ok(imagenes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener imágenes de la reparación {NumRepa}", numRepa);
                return StatusCode(500, "Error al obtener las imágenes");
            }
        }

        /// <summary>
        /// Sube una nueva imagen para una reparación
        /// </summary>
        [HttpPost("upload")]
        public async Task<ActionResult<GalimRep>> UploadImagen([FromForm] IFormFile file, [FromForm] int numRepa, [FromForm] string? descripcion)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest("No se proporcionó ningún archivo");
                }

                // Validar extensión
                var extensionesPermitidas = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp" };
                var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
                
                if (!extensionesPermitidas.Contains(extension))
                {
                    return BadRequest("Tipo de archivo no permitido. Solo se permiten imágenes.");
                }

                // Obtener el siguiente número de orden
                var maxNorit = await _context.GalimReps
                    .Where(g => g.NumRepa == numRepa)
                    .MaxAsync(g => (short?)g.Norit_Imag) ?? -1;

                short nuevoNorit = (short)(maxNorit + 1);

                // Crear nombre único para el archivo
                var nombreArchivo = $"Repa_{numRepa}_Img_{nuevoNorit}_{Guid.NewGuid()}{extension}";
                
                // Guardar archivo usando FileStorageService
                var (filePath, fileUrl) = await _fileStorage.SaveFileAsync(file, "RepaImagenes", nombreArchivo);

                // Crear registro en base de datos
                // Los campos de auditoría (AddRecord, LastUpdate, NumUsuar, CodiEmprNet) 
                // se asignan automáticamente por AuditInterceptor
                var galimRep = new GalimRep
                {
                    CodiEmpr = 0,
                    NumRepa = numRepa,
                    Norit_Imag = nuevoNorit,
                    Nom_Imag = nombreArchivo,
                    RutaImagen = fileUrl,
                    Descripcion = descripcion
                };

                _context.GalimReps.Add(galimRep);
                await _context.SaveChangesAsync();

                return Ok(galimRep);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al subir imagen para reparación {NumRepa}", numRepa);
                return StatusCode(500, "Error al subir la imagen");
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
                var galimRep = await _context.GalimReps.FindAsync(id);
                
                if (galimRep == null)
                {
                    return NotFound("Imagen no encontrada");
                }

                // Actualizar descripción
                // Los campos de auditoría (LastUpdate, NumUsuar, CodiEmprNet) 
                // se actualizan automáticamente por AuditInterceptor
                galimRep.Descripcion = descripcion;
                
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
                var galimRep = await _context.GalimReps.FindAsync(id);
                
                if (galimRep == null)
                {
                    return NotFound("Imagen no encontrada");
                }

                // Eliminar archivo físico
                if (!string.IsNullOrEmpty(galimRep.Nom_Imag))
                {
                    _fileStorage.DeleteFile("RepaImagenes", galimRep.Nom_Imag);
                }

                _context.GalimReps.Remove(galimRep);
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
