using Microsoft.Extensions.Options;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio centralizado para gestión de almacenamiento de archivos
    /// Permite configurar rutas diferentes por ambiente (Development, Production)
    /// </summary>
    public class FileStorageService
    {
        private readonly FileStorageOptions _options;
        private readonly ILogger<FileStorageService> _logger;

        public FileStorageService(IOptions<FileStorageOptions> options, ILogger<FileStorageService> logger)
        {
            _options = options.Value;
            _logger = logger;

            // Validar que la ruta base existe, si no, crearla
            if (!Directory.Exists(_options.BasePath))
            {
                _logger.LogWarning("La ruta base de almacenamiento no existe, creándola: {BasePath}", _options.BasePath);
                Directory.CreateDirectory(_options.BasePath);
            }

            _logger.LogInformation("FileStorageService inicializado con BasePath: {BasePath}", _options.BasePath);
        }

        /// <summary>
        /// Obtiene la ruta base de almacenamiento configurada
        /// </summary>
        public string BasePath => _options.BasePath;

        /// <summary>
        /// Obtiene la ruta completa para una subcarpeta específica
        /// </summary>
        /// <param name="subFolder">Nombre de la subcarpeta (ej: "usuarios", "RepaImagenes")</param>
        /// <returns>Ruta completa de la subcarpeta</returns>
        public string GetFolderPath(string subFolder)
        {
            var path = Path.Combine(_options.BasePath, subFolder);
            
            // Crear la carpeta si no existe
            if (!Directory.Exists(path))
            {
                _logger.LogInformation("Creando carpeta: {Path}", path);
                Directory.CreateDirectory(path);
            }

            return path;
        }

        /// <summary>
        /// Obtiene la ruta completa de un archivo
        /// </summary>
        /// <param name="subFolder">Subcarpeta donde se guardará el archivo</param>
        /// <param name="fileName">Nombre del archivo</param>
        /// <returns>Ruta completa del archivo</returns>
        public string GetFilePath(string subFolder, string fileName)
        {
            var folderPath = GetFolderPath(subFolder);
            return Path.Combine(folderPath, fileName);
        }

        /// <summary>
        /// Obtiene la URL relativa para acceder al archivo desde el navegador
        /// </summary>
        /// <param name="subFolder">Subcarpeta del archivo</param>
        /// <param name="fileName">Nombre del archivo</param>
        /// <returns>URL relativa (ej: "/uploads/usuarios/foto.jpg")</returns>
        public string GetFileUrl(string subFolder, string fileName)
        {
            return $"/uploads/{subFolder}/{fileName}";
        }

        /// <summary>
        /// Guarda un archivo en el almacenamiento
        /// </summary>
        /// <param name="file">Archivo a guardar</param>
        /// <param name="subFolder">Subcarpeta destino</param>
        /// <param name="fileName">Nombre del archivo (opcional, se genera uno si no se proporciona)</param>
        /// <returns>Información del archivo guardado (ruta física y URL)</returns>
        public async Task<(string FilePath, string FileUrl)> SaveFileAsync(IFormFile file, string subFolder, string? fileName = null)
        {
            if (file == null || file.Length == 0)
            {
                throw new ArgumentException("El archivo está vacío o es nulo", nameof(file));
            }

            // Generar nombre de archivo si no se proporciona
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = $"{Guid.NewGuid()}{Path.GetExtension(file.FileName)}";
            }

            var filePath = GetFilePath(subFolder, fileName);
            var fileUrl = GetFileUrl(subFolder, fileName);

            // Guardar el archivo
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            _logger.LogInformation("Archivo guardado: {FilePath}", filePath);

            return (filePath, fileUrl);
        }

        /// <summary>
        /// Elimina un archivo del almacenamiento
        /// </summary>
        /// <param name="subFolder">Subcarpeta del archivo</param>
        /// <param name="fileName">Nombre del archivo</param>
        /// <returns>True si se eliminó correctamente, False si no existía</returns>
        public bool DeleteFile(string subFolder, string fileName)
        {
            var filePath = GetFilePath(subFolder, fileName);

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                _logger.LogInformation("Archivo eliminado: {FilePath}", filePath);
                return true;
            }

            _logger.LogWarning("Intento de eliminar archivo inexistente: {FilePath}", filePath);
            return false;
        }

        /// <summary>
        /// Elimina un archivo usando su URL
        /// </summary>
        /// <param name="fileUrl">URL del archivo (ej: "/uploads/usuarios/foto.jpg")</param>
        /// <returns>True si se eliminó correctamente, False si no existía</returns>
        public bool DeleteFileByUrl(string fileUrl)
        {
            if (string.IsNullOrEmpty(fileUrl))
            {
                return false;
            }

            // Extraer subcarpeta y nombre de archivo de la URL
            // Formato esperado: /uploads/subcarpeta/archivo.ext
            var parts = fileUrl.Split('/', StringSplitOptions.RemoveEmptyEntries);
            
            if (parts.Length < 3 || parts[0] != "uploads")
            {
                _logger.LogWarning("URL de archivo inválida: {FileUrl}", fileUrl);
                return false;
            }

            var subFolder = parts[1];
            var fileName = parts[2];

            return DeleteFile(subFolder, fileName);
        }

        /// <summary>
        /// Verifica si un archivo existe
        /// </summary>
        /// <param name="subFolder">Subcarpeta del archivo</param>
        /// <param name="fileName">Nombre del archivo</param>
        /// <returns>True si existe, False si no</returns>
        public bool FileExists(string subFolder, string fileName)
        {
            var filePath = GetFilePath(subFolder, fileName);
            return File.Exists(filePath);
        }
    }

    /// <summary>
    /// Opciones de configuración para el almacenamiento de archivos
    /// </summary>
    public class FileStorageOptions
    {
        public const string SectionName = "FileStorage";

        /// <summary>
        /// Ruta base donde se almacenarán todos los archivos
        /// Ejemplo: "D:\IMPLANPROD_FILES\uploads" (Producción)
        /// Ejemplo: "C:\IMPLANPROD_FILES\uploads" (Desarrollo)
        /// </summary>
        public string BasePath { get; set; } = string.Empty;

        /// <summary>
        /// Indica si se deben servir archivos estáticos desde esta ruta
        /// </summary>
        public bool ServeStaticFiles { get; set; } = true;
    }
}
