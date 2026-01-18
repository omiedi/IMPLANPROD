using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    public class MigrationLegacyService
    {
        private readonly DataContext _context;
        private readonly ILogger<MigrationLegacyService> _logger;
        private readonly IConfiguration _configuration;

        public MigrationLegacyService(DataContext context, ILogger<MigrationLegacyService> logger, IConfiguration configuration)
        {
            _context = context;
            _logger = logger;
            _configuration = configuration;
        }

        public async Task<int> MigrateIndirepAsync(string? basePath = null)
        {
            try
            {
                if (await _context.Indireps.AnyAsync())
                {
                    _logger.LogInformation("La tabla Indirep no está vacía. Se omite la migración.");
                    return 0;
                }

                string path = basePath 
                    ?? _configuration["LegacyDataPath:BasePath"] 
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "INDIREP.DAT");
                
                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // El archivo DAT tiene registros de longitud fija sin saltos de línea
                // Cada registro tiene 128 caracteres totales:
                // - CODIGO: posiciones 1-16 (16 caracteres)
                // - DESCRIPCION: posiciones 17-80 (64 caracteres)
                // - RELLENO: posiciones 81-128 (48 caracteres de espacios/padding)
                const int RECORD_LENGTH = 128;
                const int CODIGO_LENGTH = 16;
                const int DESCRIPCION_START = 16;
                const int DESCRIPCION_LENGTH = 64;

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                int count = 0;
                int duplicatesSkipped = 0;
                var entities = new List<Indirep>();
                var processedCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Procesar el archivo en bloques de tamaño fijo
                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    // Verificar que hay suficientes bytes para un registro completo
                    int remainingBytes = fileBytes.Length - i;
                    if (remainingBytes < CODIGO_LENGTH) break; // No hay suficientes datos para un código

                    // Leer el registro completo o lo que quede
                    int bytesToRead = Math.Min(RECORD_LENGTH, remainingBytes);
                    // Usar codificación Windows-1252 (Latin-1) para caracteres acentuados del español
                    string record = System.Text.Encoding.GetEncoding(1252).GetString(fileBytes, i, bytesToRead);

                    // Extraer código (primeros 16 caracteres)
                    string codigo = record.Substring(0, CODIGO_LENGTH).Trim();
                    
                    // Extraer descripción (resto del registro)
                    string descripcion = "";
                    if (record.Length > DESCRIPCION_START)
                    {
                        int descLength = Math.Min(DESCRIPCION_LENGTH, record.Length - DESCRIPCION_START);
                        descripcion = record.Substring(DESCRIPCION_START, descLength).Trim();
                    }

                    // Validar que el código no esté vacío
                    if (string.IsNullOrEmpty(codigo)) continue;

                    // Verificar si el código ya fue procesado (duplicado)
                    if (processedCodes.Contains(codigo))
                    {
                        duplicatesSkipped++;
                        _logger.LogWarning($"Código duplicado omitido: {codigo}");
                        continue;
                    }

                    var entity = new Indirep
                    {
                        Codigo = codigo,
                        Descripcion = descripcion,
                        Status = 0
                        // AddRecord, LastUpdate, CodiEmprNet y NumUsuar se asignarán automáticamente por el AuditInterceptor
                    };
                    entities.Add(entity);
                    processedCodes.Add(codigo);
                    count++;
                }

                if (entities.Any())
                {
                    await _context.Indireps.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Se migraron {count} registros a la tabla Indirep. Duplicados omitidos: {duplicatesSkipped}");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Indirep");
                throw;
            }
        }

        public async Task<int> MigrateRepuesreAsync(string? basePath = null)
        {
            try
            {
                if (await _context.Repuesres.AnyAsync())
                {
                    _logger.LogInformation("La tabla Repuesre no está vacía. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "REPUESRE.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // El archivo REPUESRE tiene registros de longitud fija de 64 bytes:
                // - CODIGOREPARACION: posiciones 1-16 (16 caracteres)
                // - CODIGOINTERNO: posiciones 17-18 (2 bytes, Int16 - CVI)
                // - USO: posiciones 19-22 (4 bytes, Single/float en formato MBF - CVS)
                // - RELLENO: posiciones 23-64 (42 bytes de padding)
                const int RECORD_LENGTH = 64;
                const int CODIGO_REPARACION_LENGTH = 16;
                const int CODIGO_INTERNO_START = 16;
                const int USO_START = 18;

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                int count = 0;
                int duplicatesSkipped = 0;
                var entities = new List<Repuesre>();
                var processedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Procesar el archivo en bloques de tamaño fijo
                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    // Verificar que hay suficientes bytes para un registro completo
                    int remainingBytes = fileBytes.Length - i;
                    if (remainingBytes < RECORD_LENGTH) break;

                    // Extraer código de reparación (primeros 16 caracteres)
                    string codigoReparacion = System.Text.Encoding.GetEncoding(1252)
                        .GetString(fileBytes, i, CODIGO_REPARACION_LENGTH).Trim();

                    // Validar que el código no esté vacío
                    if (string.IsNullOrEmpty(codigoReparacion)) continue;

                    // Extraer CodigoInterno (2 bytes, Int16) - equivalente a CVI en VB6
                    short codigoInterno = BitConverter.ToInt16(fileBytes, i + CODIGO_INTERNO_START);

                    // Extraer Uso (4 bytes) - formato Microsoft Binary Format (MBF)
                    // Necesitamos usar la función CVS personalizada
                    byte[] usoBytes = new byte[4];
                    Array.Copy(fileBytes, i + USO_START, usoBytes, 0, 4);
                    float uso = ConvertMBFToFloat(usoBytes);

                    // Crear clave única para detectar duplicados
                    string key = $"{codigoReparacion}|{codigoInterno}";

                    // Verificar si la combinación ya fue procesada (duplicado)
                    if (processedKeys.Contains(key))
                    {
                        duplicatesSkipped++;
                        _logger.LogWarning($"Registro duplicado omitido: {codigoReparacion} - {codigoInterno}");
                        continue;
                    }

                    var entity = new Repuesre
                    {
                        CodigoReparacion = codigoReparacion,
                        CodigoInterno = codigoInterno,
                        Uso = (decimal)uso,
                        Status = 0
                        // AddRecord, LastUpdate, CodiEmprNet y NumUsuar se asignarán automáticamente por el AuditInterceptor
                    };
                    entities.Add(entity);
                    processedKeys.Add(key);
                    count++;
                }

                if (entities.Any())
                {
                    await _context.Repuesres.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Se migraron {count} registros a la tabla Repuesre. Duplicados omitidos: {duplicatesSkipped}");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Repuesre");
                throw;
            }
        }

        public async Task<int> MigrateSecoprepAsync(string? basePath = null)
        {
            try
            {
                if (await _context.Secopreps.AnyAsync())
                {
                    _logger.LogInformation("La tabla Secoprep no está vacía. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "SECOPREP.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // El archivo SECOPREP tiene registros de longitud fija de 64 bytes:
                // - CODIGOREPARACION: posiciones 1-16 (16 caracteres)
                // - NUMEROOPERACION: posiciones 17-18 (2 bytes, Int16 - CVI)
                // - DESCRIPCION: posiciones 19-66 (48 caracteres)
                // Total: 66 bytes, pero el archivo usa bloques de 64, así que verificaremos
                const int RECORD_LENGTH = 128;
                const int CODIGO_REPARACION_LENGTH = 16;
                const int NUMERO_OPERACION_START = 16;
                const int DESCRIPCION_START = 18;
                const int DESCRIPCION_LENGTH = 48;

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                int count = 0;
                int duplicatesSkipped = 0;
                var entities = new List<Secoprep>();
                var processedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Procesar el archivo en bloques de tamaño fijo
                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    // Verificar que hay suficientes bytes para un registro completo
                    int remainingBytes = fileBytes.Length - i;
                    if (remainingBytes < RECORD_LENGTH) break;

                    // Extraer código de reparación (primeros 16 caracteres)
                    string codigoReparacion = System.Text.Encoding.GetEncoding(1252)
                        .GetString(fileBytes, i, CODIGO_REPARACION_LENGTH).Trim();

                    // Validar que el código no esté vacío
                    if (string.IsNullOrEmpty(codigoReparacion)) continue;

                    // Extraer NumeroOperacion (2 bytes, Int16) - equivalente a CVI en VB6
                    short numeroOperacion = BitConverter.ToInt16(fileBytes, i + NUMERO_OPERACION_START);

                    // Extraer Descripción (48 caracteres) pos 18-66
                    string descripcion = "";
                    int bytesDisponibles = Math.Min(DESCRIPCION_LENGTH, remainingBytes - DESCRIPCION_START);
                    if (bytesDisponibles > 0)
                    {
                        descripcion = System.Text.Encoding.GetEncoding(1252)
                            .GetString(fileBytes, i + DESCRIPCION_START, bytesDisponibles).Trim();
                    }

                    // --- NUEVOS CAMPOS ---
                    // Definiendo offsets relativos al final de la descripción (Offset 66)
                    const int EQUIPO_START = 64; // 18 + 48
                    const int EQUIPO_LENGTH = 8;
                    const int HORAS_START = 72; // 64 + 8
                    const int OPERARIOS_START = 98; // 72 + 4 + 22 (Relleno)
                    const int OPERARIOS_LENGTH = 4;

                    // Equipo (8 caracteres) -> Texto (Máximo 6 según requerimiento)
                    string? equipo = null;
                    if (remainingBytes >= EQUIPO_START + EQUIPO_LENGTH)
                    {
                        string equipoStr = System.Text.Encoding.GetEncoding(1252)
                             .GetString(fileBytes, i + EQUIPO_START, EQUIPO_LENGTH).Trim();
                        
                        if (!string.IsNullOrEmpty(equipoStr))
                        {
                            // Tomar hasta 6 caracteres
                            equipo = equipoStr.Length > 6 ? equipoStr.Substring(0, 6) : equipoStr;
                        }
                    }

                    // horas (4 bytes MBF)
                    double horas = 0;
                    if (remainingBytes >= HORAS_START + 4)
                    {
                        byte[] horasBytes = new byte[4];
                        Array.Copy(fileBytes, i + HORAS_START, horasBytes, 0, 4);
                        horas = ConvertMBFToFloat(horasBytes);
                    }

                    // Operarios (4 bytes MBF) -> Int
                    int operarios = 0;
                    if (remainingBytes >= OPERARIOS_START + OPERARIOS_LENGTH)
                    {
                        byte[] operariosBytes = new byte[OPERARIOS_LENGTH];
                        Array.Copy(fileBytes, i + OPERARIOS_START, operariosBytes, 0, OPERARIOS_LENGTH);
                        float opFloat = ConvertMBFToFloat(operariosBytes);
                        operarios = (int)opFloat;
                    }

                    // Categoría (Sin posición definida aún - Default 0)
                    int categoria = 0;

                    // Crear clave única para detectar duplicados
                    string key = $"{codigoReparacion}|{numeroOperacion}";

                    // Verificar si la combinación ya fue procesada (duplicado)
                    if (processedKeys.Contains(key))
                    {
                        duplicatesSkipped++;
                        _logger.LogWarning($"Registro duplicado omitido: {codigoReparacion} - {numeroOperacion}");
                        continue;
                    }

                    var entity = new Secoprep
                    {
                        CodigoReparacion = codigoReparacion,
                        NumeroOperacion = numeroOperacion,
                        Descripcion = descripcion,
                        Status = 0,
                        Horas = horas,
                        Equipo = equipo,
                        CantidadOperarios = operarios,
                        Categoria = categoria
                        // AddRecord, LastUpdate, CodiEmprNet y NumUsuar se asignarán automáticamente por el AuditInterceptor
                    };
                    entities.Add(entity);
                    processedKeys.Add(key);
                    count++;
                }

                if (entities.Any())
                {
                    await _context.Secopreps.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Se migraron {count} registros a la tabla Secoprep. Duplicados omitidos: {duplicatesSkipped}");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Secoprep");
                throw;
            }
        }

        public async Task<int> MigrateObsoprepAsync(string? basePath = null)
        {
            try
            {
                // Verificar si ya se procesaron las observaciones
                // Contamos cuántos registros en Secoprep tienen observaciones
                var registrosConObservaciones = await _context.Secopreps
                    .Where(s => !string.IsNullOrEmpty(s.Observacion))
                    .CountAsync();
                
                if (registrosConObservaciones > 0)
                {
                    _logger.LogInformation($"Ya existen {registrosConObservaciones} registros con observaciones en Secoprep. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "OBSOPREP.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // El archivo OBSOPREP tiene registros de longitud fija de 256 bytes:
                // - CODIGOREPARACION: posiciones 1-16 (16 caracteres)
                // - NUMEROOPERACION: posiciones 17-18 (2 bytes, Int16 - CVI)
                // - DESCRIPCION (Observacion): posiciones 19-256 (238 caracteres)
                const int RECORD_LENGTH = 256;
                const int CODIGO_REPARACION_LENGTH = 16;
                const int NUMERO_OPERACION_START = 16;
                const int DESCRIPCION_START = 18;
                const int DESCRIPCION_LENGTH = 238; // 256 - 18 = 238 bytes desde posición 19 hasta el final

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                int count = 0;
                int notFoundCount = 0;

                // Procesar el archivo en bloques de tamaño fijo
                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    // Verificar que hay suficientes bytes para un registro completo
                    int remainingBytes = fileBytes.Length - i;
                    if (remainingBytes < RECORD_LENGTH) break;

                    // Extraer código de reparación (primeros 16 caracteres)
                    string codigoReparacion = System.Text.Encoding.GetEncoding(1252)
                        .GetString(fileBytes, i, CODIGO_REPARACION_LENGTH).Trim();

                    // Validar que el código no esté vacío
                    if (string.IsNullOrEmpty(codigoReparacion)) continue;

                    // Extraer NumeroOperacion (2 bytes, Int16) - equivalente a CVI en VB6
                    short numeroOperacion = BitConverter.ToInt16(fileBytes, i + NUMERO_OPERACION_START);

                    // Extraer Descripcion/Observacion (128 caracteres)
                    string observacion = "";
                    int bytesDisponibles = Math.Min(DESCRIPCION_LENGTH, remainingBytes - DESCRIPCION_START);
                    if (bytesDisponibles > 0)
                    {
                        observacion = System.Text.Encoding.GetEncoding(1252)
                            .GetString(fileBytes, i + DESCRIPCION_START, bytesDisponibles).Trim();
                    }

                    // Buscar el registro correspondiente en SECOPREP y actualizar
                    var secoprepRecord = await _context.Secopreps
                        .FirstOrDefaultAsync(s => s.CodigoReparacion == codigoReparacion 
                                                && s.NumeroOperacion == numeroOperacion);

                    if (secoprepRecord != null)
                    {
                        secoprepRecord.Observacion = observacion;
                        count++;
                    }
                    else
                    {
                        notFoundCount++;
                        _logger.LogWarning($"No se encontró registro en SECOPREP para: {codigoReparacion} - {numeroOperacion}");
                    }
                }

                if (count > 0)
                {
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Se actualizaron {count} registros en SECOPREP con observaciones. No encontrados: {notFoundCount}");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Obsoprep");
                throw;
            }
        }

        public async Task<int> MigrateDescrepAsync(string? basePath = null)
        {
            try
            {
                if (await _context.Descreps.AnyAsync())
                {
                    _logger.LogInformation("La tabla Descrep no está vacía. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "DESCREP.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    return 0;
                }

                // El archivo DESCREP.DAT tiene registros de longitud fija 80
                // CODIGO: posiciones 1-16 (16 caracteres)
                // ESPECIFICACION: posiciones 17-80 (64 caracteres)
                const int RECORD_LENGTH = 80;
                const int CODIGO_LENGTH = 16;
                const int ESPECIFICACION_START = 16;
                const int ESPECIFICACION_LENGTH = 64;

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                
                // Diccionario para concatenar especificaciones por código
                var descripcionesMap = new Dictionary<string, System.Text.StringBuilder>(StringComparer.OrdinalIgnoreCase);

                // Procesar el archivo
                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    int remainingBytes = fileBytes.Length - i;
                    if (remainingBytes < RECORD_LENGTH) break;

                    // Extraer código (posiciones 1-16)
                    string codigo = System.Text.Encoding.GetEncoding(1252)
                        .GetString(fileBytes, i, CODIGO_LENGTH).Trim();

                    if (string.IsNullOrEmpty(codigo)) continue;

                    // Extraer especificación (posiciones 17-80, es decir 64 caracteres)
                    string especificacion = "";
                    int bytesDisponibles = Math.Min(ESPECIFICACION_LENGTH, remainingBytes - ESPECIFICACION_START);
                    if (bytesDisponibles > 0)
                    {
                        // Extraer bytes sin convertir primero
                        byte[] especBytes = new byte[bytesDisponibles];
                        Array.Copy(fileBytes, i + ESPECIFICACION_START, especBytes, 0, bytesDisponibles);
                        
                        especificacion = System.Text.Encoding.GetEncoding(1252).GetString(especBytes);
                        
                        // RTrim solo de espacios (mantener CR/LF internos)
                        // Usar un bucle manual para solo quitar espacios del final
                        int endIndex = especificacion.Length - 1;
                        while (endIndex >= 0 && especificacion[endIndex] == ' ')
                        {
                            endIndex--;
                        }
                        if (endIndex < especificacion.Length - 1)
                        {
                            especificacion = especificacion.Substring(0, endIndex + 1);
                        }
                    }

                    // Inicializar StringBuilder si es la primera vez que vemos este código
                    if (!descripcionesMap.ContainsKey(codigo))
                    {
                        descripcionesMap[codigo] = new System.Text.StringBuilder();
                    }
                    else
                    {
                        // Si ya tiene contenido, agregar \r\n entre registros
                        if (descripcionesMap[codigo].Length > 0)
                        {
                            descripcionesMap[codigo].Append("\r\n");
                        }
                    }

                    // Agregar la especificación (incluso si está vacía, para mantener estructura)
                    if (!string.IsNullOrEmpty(especificacion))
                    {
                        descripcionesMap[codigo].Append(especificacion);
                    }
                }

                // Crear entidades
                var entities = new List<Descrep>();

                foreach (var kvp in descripcionesMap)
                {
                    // No hacer TrimEnd() final para no borrar ENTER finales
                    string especificacionFinal = kvp.Value.ToString();
                    
                    // Log para debugging - mostrar caracteres especiales
                    var caracteresEspeciales = especificacionFinal.Count(c => c == '\r' || c == '\n');
                    if (caracteresEspeciales > 0)
                    {
                        _logger.LogInformation($"Código {kvp.Key}: {caracteresEspeciales} caracteres CR/LF encontrados");
                    }
                    
                    entities.Add(new Descrep
                    {
                        CodigoReparacion = kvp.Key,
                        Especificacion = especificacionFinal,
                        Status = 0,
                    });
                }

                if (entities.Any())
                {
                    await _context.Descreps.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Se migraron {entities.Count} registros a la tabla Descrep.");
                return entities.Count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Descrep");
                throw;
            }
        }


        /// <summary>
        /// Convierte un número en formato Microsoft Binary Format (MBF) a IEEE float
        /// Equivalente a la función CVS de VB6
        /// </summary>
        private float ConvertMBFToFloat(byte[] mbfBytes)
        {
            if (mbfBytes == null || mbfBytes.Length != 4)
                return 0f;

            byte aa1 = mbfBytes[0];
            byte aa2 = mbfBytes[1];
            byte aa3 = mbfBytes[2];
            byte aa4 = mbfBytes[3];

            // Verificar si todos los bytes son cero o si el exponente es 0
            if (aa4 == 0 || (aa1 == 0 && aa2 == 0 && aa3 == 0 && aa4 == 0))
                return 0f;

            // Calcular el signo basado en el bit más significativo de aa3
            int signo = (aa3 & 0x80) == 0 ? 1 : -1;

            // Calcular el exponente (aa4 - 129 para MBF)
            int exponent = aa4 - 129;
            
            // Calcular la mantisa
            // Remover el bit de signo de aa3 para obtener solo los bits de mantisa
            int a31 = aa3 & 0x7F; // Máscara 0x7F = 01111111 para quitar el bit de signo
            
            // La mantisa en MBF NO incluye el 1.0 implícito (a diferencia de IEEE)
            // Mantisa = (a31/128 + aa2/32768 + aa1/8388608)
            double mantisa = (a31 / 128.0) + (aa2 / 32768.0) + (aa1 / 8388608.0);
            
            // Resultado = signo * (1 + mantisa) * 2^exponente
            double result = signo * (1.0 + mantisa) * Math.Pow(2, exponent);
            
            return (float)result;
        }
    }
}
