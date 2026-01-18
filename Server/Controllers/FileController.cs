using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Net;

using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Authorization;

namespace IMPLANPROD.Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly FileStorageService _fileStorage;

        public FileController(DataContext context, FileStorageService fileStorage)
        {
            _context = context;
            _fileStorage = fileStorage;
        }

        [HttpPost]
        public async Task<ActionResult<List<UploadMaster>>> UploadFile(List<IFormFile> files, [FromQuery] int codint = 0)
        {
            List<UploadMaster> uploadResults = new List<UploadMaster>();

            foreach (var file in files)
            {
                var uploadResult = new UploadMaster();
                string trustedFileNameForFileStorage;
                var untrustedFileName = file.FileName;
                uploadResult.Nombre = untrustedFileName;
                //var trustedFileNameForDisplay = WebUtility.HtmlEncode(untrustedFileName);

                trustedFileNameForFileStorage = Path.GetRandomFileName();
                var (filePath, fileUrl) = await _fileStorage.SaveFileAsync(file, "", trustedFileNameForFileStorage);
                uploadResult.Nombre = file.FileName;
                uploadResult.StoredFileName = trustedFileNameForFileStorage;
                uploadResult.ContentType = file.ContentType;
                uploadResult.Codint = codint;
                uploadResult.FechaSubida = DateTime.Now;
                uploadResults.Add(uploadResult);

                try
                {
                _context.UploadMaster.Add(uploadResult);
                await _context.SaveChangesAsync();

                }
                catch (Exception ex)
                {

                    throw;
                }
            }

            await _context.SaveChangesAsync();

            return Ok(uploadResults);
        }

        [HttpGet]
        [Route("FileList")]
        public async Task<ActionResult<List<UploadMaster>>> GetEjemplo()
        {
            var lista = await _context.UploadMaster.ToListAsync();
            return Ok(lista);
        }
        [HttpGet]
        //retorna archivos adjuntos segun codigo interno
        [Route("FileList/{codint:int}")]
        public async Task<ActionResult<List<UploadMaster>>> GetEjemplo(int codint)
        {
            var lista = await _context.UploadMaster
            .Where(a => a.Codint == codint)
            .OrderByDescending(a => a.FechaSubida)

                .ToListAsync();

            if (lista == null || !lista.Any())
            {
                return NotFound(new { message = "No se encontraron archivos para el código activo." });
            }
            return Ok(lista);
        }

        [HttpPost]
        [Route("getFile")]
        public async Task<IActionResult> ObtenerFile(UploadMaster result)
        {
            //var uploadResult = await _context.UploadResults.FirstOrDefaultAsync(u => u.StoredFileName.Equals(fileName));

            var path = _fileStorage.GetFilePath("", result.StoredFileName);

            var memory = new MemoryStream();
            using (var stream = new FileStream(path, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;
            //return File(memory, "image/png", Path.GetFileName(path));//para imagenes pnd
            return File(memory.ToArray(), result.ContentType, Path.GetFileName(path));//Para excel

        }

        [HttpDelete]
        [Route("DeleteFile/{id:int}")]
        public async Task<IActionResult> DeleteFile(int id)
        {
            var archivo = await _context.UploadMaster.FindAsync(id);
            if (archivo == null)
            {
                return NotFound(new { message = "Archivo no encontrado." });
            }

            if (!string.IsNullOrEmpty(archivo!.StoredFileName))
            {
                _fileStorage.DeleteFile("", archivo.StoredFileName);
            }

            _context.UploadMaster.Remove(archivo);
            await _context.SaveChangesAsync();

            return Ok(new { message = "Archivo eliminado exitosamente." });
        }
    }


}


