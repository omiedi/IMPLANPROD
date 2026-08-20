using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
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
        /// Migra los datos desde el archivo plano CODIPROC.DAT a la tabla SQL CODIPROC
        /// 
        /// ESTRUCTURA DEL ARCHIVO PLANO (512 bytes por registro):
        /// - Offset  0 ( 2 bytes): CodiEmpr     → Int16
        /// - Offset  2 ( 2 bytes): CodProc      → Int16  (clave de negocio: Ci% en VB6)
        /// - Offset  4 (12 bytes): (no mapeados en el formulario PLAINSP07)
        /// - Offset 16 ( 2 bytes): AQL          → Int16  (Nivel de Calidad Aceptable)
        /// - Offset 18 ( 2 bytes): NivelPct     → Int16  (Nivel porcentual)
        /// - Offset 20 (32 bytes): FrecInspeccion → char(32)
        /// - Offset 52 (48 bytes): PuntoDetencion → char(48)
        /// - Offset100 (24 bytes): NombrePreparo  → char(24)
        /// - Offset124 ( 2 bytes): FechaPreparo   → Int16 (fecha OLE/VB6: días desde 30/12/1899)
        /// - Offset126 (24 bytes): NombreAprobo   → char(24)
        /// - Offset150 ( 2 bytes): FechaAprobo    → Int16 (fecha OLE/VB6)
        /// - Offset152 ( 8 bytes): (no mapeados en el formulario PLAINSP07)
        /// - Offset160 (352 bytes): Observaciones → char(352)
        /// 
        /// CONDICIÓN DE SALTO: si CodProc == 0, el registro está vacío o borrado → se omite
        /// DUPLICADOS: se detectan por CodProc y se omiten
        /// </summary>
        public async Task<int> MigrateCodiprocAsync(string? basePath = null)
        {
            try
            {
                if (await _context.Codiprocs.AnyAsync())
                {
                    _logger.LogInformation("La tabla Codiproc no está vacía. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "CODIPROC.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // Tamaño fijo del registro: 512 bytes (igual al tamaño total del registro VB6)
                const int RECORD_LENGTH = 512;

                // Offsets (base 0) derivados del código VB6:
                // Mid$(YY$, N, L) en VB6 (base 1) → offset en C# = N - 1
                const int COD_INTE_OFFSET        = 0;   // bytes 1-2  → Cod_Inte (CVI: 256*byte2 + byte1 = BitConverter.ToInt16)
                const int CODPROC_OFFSET         = 2;   // bytes 3-4 (clave Ci%)
                const int AQL_OFFSET             = 16;  // bytes 17-18
                const int NIVELPCT_OFFSET        = 18;  // bytes 19-20
                const int FRECINSPECCION_OFFSET  = 20;  // bytes 21-52 (32 bytes)
                const int FRECINSPECCION_LENGTH  = 32;
                const int PUNTODET_OFFSET        = 52;  // bytes 53-100 (48 bytes)
                const int PUNTODET_LENGTH        = 48;
                const int NOMBREPREPARO_OFFSET   = 100; // bytes 101-124 (24 bytes)
                const int NOMBREPREPARO_LENGTH   = 24;
                const int FECHAPREPARO_OFFSET    = 124; // bytes 125-126
                const int NOMBREAPROBO_OFFSET    = 126; // bytes 127-150 (24 bytes)
                const int NOMBREAPROBO_LENGTH    = 24;
                const int FECHAAPROBO_OFFSET     = 150; // bytes 151-152
                const int OBSERVACIONES_OFFSET   = 160; // bytes 161-512 (352 bytes)
                const int OBSERVACIONES_LENGTH   = 352;

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                int count = 0;
                int duplicatesSkipped = 0;
                int emptySkipped = 0;
                var entities = new List<Codiproc>();
                var processedCodes = new HashSet<int>();

                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    // Verificar que hay suficientes bytes para un registro completo
                    if (fileBytes.Length - i < RECORD_LENGTH) break;

                    // Leer Cod_Inte (offset 0-1): es el campo identificador real del registro.
                    // NOTA: CodProc (offset 2-3) es SIEMPRE 0 en el archivo CODIPROC.DAT.
                    //       El campo con valores no-nulos es Cod_Inte (offset 0-1).
                    short codInte = BitConverter.ToInt16(fileBytes, i + COD_INTE_OFFSET);

                    // Si Cod_Inte == 0 el registro está vacío o fue borrado → omitir
                    if (codInte == 0)
                    {
                        emptySkipped++;
                        continue;
                    }

                    // Verificar duplicados por Cod_Inte (campo identificador real)
                    if (processedCodes.Contains(codInte))
                    {
                        duplicatesSkipped++;
                        _logger.LogWarning($"Cod_Inte duplicado omitido: {codInte}");
                        continue;
                    }

                    // Leer CodProc (posiciones 3-4, CVI) - en el archivo real es siempre 0
                    short codProc = BitConverter.ToInt16(fileBytes, i + CODPROC_OFFSET);

                    // Leer AQL y NivelPct (CVI en VB6 → Int16)
                    short aql      = BitConverter.ToInt16(fileBytes, i + AQL_OFFSET);
                    short nivelPct = BitConverter.ToInt16(fileBytes, i + NIVELPCT_OFFSET);

                    // Leer campos de texto (Windows-1252 para caracteres acentuados)
                    string frecInspeccion = System.Text.Encoding.GetEncoding(1252)
                        .GetString(fileBytes, i + FRECINSPECCION_OFFSET, FRECINSPECCION_LENGTH).Trim();

                    string puntoDetencion = System.Text.Encoding.GetEncoding(1252)
                        .GetString(fileBytes, i + PUNTODET_OFFSET, PUNTODET_LENGTH).Trim();

                    string nombrePreparo = System.Text.Encoding.GetEncoding(1252)
                        .GetString(fileBytes, i + NOMBREPREPARO_OFFSET, NOMBREPREPARO_LENGTH).Trim();

                    string nombreAprobo = System.Text.Encoding.GetEncoding(1252)
                        .GetString(fileBytes, i + NOMBREAPROBO_OFFSET, NOMBREAPROBO_LENGTH).Trim();

                    string observaciones = System.Text.Encoding.GetEncoding(1252)
                        .GetString(fileBytes, i + OBSERVACIONES_OFFSET, OBSERVACIONES_LENGTH).TrimEnd();

                    // Convertir fechas VB6 con el formato FECHATEX$ (codificación custom)
                    // Formato: FEX = Día + 32*Mes + 416*(Año2d - 80) — NO es OLE Automation date
                    short fechaPreparoRaw = BitConverter.ToInt16(fileBytes, i + FECHAPREPARO_OFFSET);
                    DateTime? fechaPreparo = ConvertFechatexToDateTime(fechaPreparoRaw);

                    short fechaAproboRaw = BitConverter.ToInt16(fileBytes, i + FECHAAPROBO_OFFSET);
                    DateTime? fechaAprobo = ConvertFechatexToDateTime(fechaAproboRaw);

                    var entity = new Codiproc
                    {
                        CodiEmpr       = 0,                                           // Siempre 0 según reglas globales
                        Cod_Inte       = (int)codInte,
                        CodProc        = (int)codProc,
                        AQL            = (int)aql,
                        NivelPct       = (int)nivelPct,
                        FrecInspeccion = string.IsNullOrEmpty(frecInspeccion) ? null : frecInspeccion,
                        PuntoDetencion = string.IsNullOrEmpty(puntoDetencion) ? null : puntoDetencion,
                        NombrePreparo  = string.IsNullOrEmpty(nombrePreparo) ? null : nombrePreparo,
                        FechaPreparo   = fechaPreparo,
                        NombreAprobo   = string.IsNullOrEmpty(nombreAprobo) ? null : nombreAprobo,
                        FechaAprobo    = fechaAprobo,
                        Observaciones  = string.IsNullOrEmpty(observaciones) ? null : observaciones
                        // AddRecord, LastUpdate, CodiEmprNet, NumUsuar → asignados por el AuditInterceptor
                    };

                    entities.Add(entity);
                    processedCodes.Add(codInte);
                    count++;
                }

                if (entities.Any())
                {
                    await _context.Codiprocs.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Migración CODIPROC completada: {count} registros migrados, {duplicatesSkipped} duplicados omitidos, {emptySkipped} registros vacíos omitidos.");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Codiproc");
                throw;
            }
        }

        /// <summary>
        /// Migra los datos desde el archivo plano PLAINSP.DAT a la tabla SQL PLAINSP
        ///
        /// ESTRUCTURA DEL ARCHIVO PLANO (128 bytes por registro):
        /// - Offset  0 ( 2 bytes): Cod_Inte     → Int16 CVI (clave hacia CODIPROC.Cod_Inte)
        /// - Offset  2 (32 bytes): DimSegunPlano → char(32)  Mid$(Y$, 3, 32)
        /// - Offset 34 (30 bytes): Instrumento   → char(30)  Mid$(Y$, 35, 30)
        /// - Offset 64 (64 bytes): Observaciones → char(64)  Mid$(Y$, 65, 64)
        ///
        /// CONDICIÓN DE SALTO: DimSegunPlano y Instrumento vacíos → registro vacío/borrado
        /// NroOrden: contador secuencial dentro de cada Cod_Inte (orden físico del archivo)
        /// </summary>
        public async Task<int> MigratePlainspsAsync(string? basePath = null)
        {
            try
            {
                if (await _context.Plainsps.AnyAsync())
                {
                    _logger.LogInformation("La tabla Plainsp no está vacía. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "PLAINSP.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // Tamaño de registro derivado del código VB6: campos hasta posición 128
                const int RECORD_LENGTH = 128;

                // Offsets (base 0) → Mid$(Y$, N, L) en VB6 (base 1) = offset N-1
                const int COD_INTE_OFFSET       = 0;    // bytes 1-2   → CVI (Cod_Inte)
                const int DIMSEGUNPLANO_OFFSET  = 2;    // bytes 3-34  → char 32
                const int DIMSEGUNPLANO_LENGTH  = 32;
                const int INSTRUMENTO_OFFSET    = 34;   // bytes 35-64 → char 30
                const int INSTRUMENTO_LENGTH    = 30;
                const int OBSERVACIONES_OFFSET  = 64;   // bytes 65-128 → char 64
                const int OBSERVACIONES_LENGTH  = 64;

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                var enc = System.Text.Encoding.GetEncoding(1252);
                var entities = new List<Plainsp>();
                var orderByCodInte = new Dictionary<int, int>();   // secuencia por Cod_Inte
                int count = 0;
                int emptySkipped = 0;

                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    if (fileBytes.Length - i < RECORD_LENGTH) break;

                    // Cod_Inte: CVI = BitConverter.ToInt16 (misma lógica que la función CVI de VB6)
                    short codInteRaw = BitConverter.ToInt16(fileBytes, i + COD_INTE_OFFSET);
                    int codInte = (int)codInteRaw;

                    // Solo se salta si Cod_Inte == 0: ranura vacía o borrada en el archivo ISAM
                    // Los registros con texto vacío pero con Cod_Inte válido se crean igual
                    // (se guardan con DimSegunPlano/Instrumento/Observaciones en null)
                    if (codInte == 0)
                    {
                        emptySkipped++;
                        continue;
                    }

                    string dimSegunPlano = enc.GetString(fileBytes, i + DIMSEGUNPLANO_OFFSET, DIMSEGUNPLANO_LENGTH).Trim();
                    string instrumento   = enc.GetString(fileBytes, i + INSTRUMENTO_OFFSET,   INSTRUMENTO_LENGTH).Trim();
                    string observaciones = enc.GetString(fileBytes, i + OBSERVACIONES_OFFSET, OBSERVACIONES_LENGTH).TrimEnd();

                    // Calcular NroOrden secuencial por Cod_Inte (refleja el orden físico del archivo)
                    if (!orderByCodInte.ContainsKey(codInte))
                        orderByCodInte[codInte] = 0;
                    orderByCodInte[codInte]++;
                    int nroOrden = orderByCodInte[codInte];

                    var entity = new Plainsp
                    {
                        CodiEmpr      = 0,                                                      // Siempre 0 según reglas globales
                        Cod_Inte      = codInte,
                        NroOrden      = nroOrden,
                        DimSegunPlano = string.IsNullOrEmpty(dimSegunPlano) ? null : dimSegunPlano,
                        Instrumento   = string.IsNullOrEmpty(instrumento)   ? null : instrumento,
                        Observaciones = string.IsNullOrEmpty(observaciones) ? null : observaciones
                        // AddRecord, LastUpdate, CodiEmprNet, NumUsuar → asignados por el AuditInterceptor
                    };

                    entities.Add(entity);
                    count++;
                }

                if (entities.Any())
                {
                    await _context.Plainsps.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Migración PLAINSP completada: {count} registros migrados, {emptySkipped} registros vacíos omitidos.");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Plainsp");
                throw;
            }
        }

        /// <summary>
        /// Migra los datos desde el archivo plano INSTRUCI.DAT a la tabla SQL INSTRUCI
        ///
        /// ESTRUCTURA DEL ARCHIVO PLANO (128 bytes por registro):
        /// - Offset   0 (16 bytes): Codigo      → char(16)  left$(label37)
        /// - Offset  16 (48 bytes): Descripcion → char(48)  AJUSTI$(Text3, 48)
        /// - Offset  64 ( 4 bytes): Revision    → char(4)   AJUSTI$(Text1, 4)
        /// - Offset  68 ( 2 bytes): Cod_Inte    → Int16 CVI AJUSTI$(LABEL28, 4) len 2
        /// - Offset  70 (24 bytes): Preparo     → char(24)  AJUSTI$(Text4(0), 24)
        /// - Offset  94 ( 2 bytes): FechaPrep   → Int16 MKI$(FECHANUM(Text5(0)))
        /// - Offset  96 (24 bytes): Aprobo      → char(24)  AJUSTI$(Text4(1), 24)
        /// - Offset 120 ( 2 bytes): FechaPro    → Int16 MKI$(FECHANUM(Text5(1)))
        /// - Offset 122 ( 6 bytes): (padding hasta 128 bytes)
        ///
        /// CONDICIÓN DE SALTO: Cod_Inte == 0 → ranura vacía o borrada
        /// DUPLICADOS: detectados por Cod_Inte
        /// </summary>
        public async Task<int> MigrateInstruciAsync(string? basePath = null)
        {
            try
            {
                if (await _context.Instrucis.AnyAsync())
                {
                    _logger.LogInformation("La tabla Instruci no está vacía. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "INSTRUCI.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // Tamaño fijo del registro: 128 bytes (3968 bytes / 31 registros)
                const int RECORD_LENGTH = 128;

                // Offsets (base 0) → VB6 REPLA(XX$, valor, pos_base1, len): offset = pos - 1
                const int CODIGO_OFFSET      = 0;    // VB6 pos 1,  len 16 → char(16)
                const int CODIGO_LENGTH      = 16;
                const int DESCRIPCION_OFFSET = 16;   // VB6 pos 17, len 48 → char(48)
                const int DESCRIPCION_LENGTH = 48;
                const int REVISION_OFFSET    = 64;   // VB6 pos 65, len 4  → char(4)
                const int REVISION_LENGTH    = 4;
                const int COD_INTE_OFFSET    = 68;   // VB6 pos 69, len 2  → Int16 CVI
                const int PREPARO_OFFSET     = 70;   // VB6 pos 71, len 24 → char(24)
                const int PREPARO_LENGTH     = 24;
                const int FECHAPREP_OFFSET   = 94;   // VB6 pos 95, len 2  → Int16 FECHANUM
                const int APROBO_OFFSET      = 96;   // VB6 pos 97, len 24 → char(24)
                const int APROBO_LENGTH      = 24;
                const int FECHAPRO_OFFSET    = 120;  // VB6 pos 121, len 2 → Int16 FECHANUM

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                var enc       = System.Text.Encoding.GetEncoding(1252);
                var entities  = new List<Instruci>();
                var processedCodes = new HashSet<int>();
                int count        = 0;
                int emptySkipped = 0;
                int duplicatesSkipped = 0;

                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    if (fileBytes.Length - i < RECORD_LENGTH) break;

                    // Cod_Inte: campo identificador real (Int16 CVI, little-endian)
                    short codInteRaw = BitConverter.ToInt16(fileBytes, i + COD_INTE_OFFSET);
                    int   codInte    = (int)codInteRaw;

                    // Ranura vacía o borrada
                    if (codInte == 0)
                    {
                        emptySkipped++;
                        continue;
                    }

                    // Detectar duplicados
                    if (processedCodes.Contains(codInte))
                    {
                        duplicatesSkipped++;
                        _logger.LogWarning($"Cod_Inte duplicado en INSTRUCI omitido: {codInte}");
                        continue;
                    }

                    string codigo      = enc.GetString(fileBytes, i + CODIGO_OFFSET,      CODIGO_LENGTH).Trim();
                    string descripcion = enc.GetString(fileBytes, i + DESCRIPCION_OFFSET, DESCRIPCION_LENGTH).Trim();
                    string revision    = enc.GetString(fileBytes, i + REVISION_OFFSET,    REVISION_LENGTH).Trim();
                    string preparo     = enc.GetString(fileBytes, i + PREPARO_OFFSET,     PREPARO_LENGTH).Trim();
                    string aprobo      = enc.GetString(fileBytes, i + APROBO_OFFSET,      APROBO_LENGTH).Trim();

                    // Fechas codificadas con FECHANUM (mismo formato FECHATEX$)
                    // FEX = Día + 32*Mes + 416*(Año2d - 80)
                    short fechaPrepRaw = BitConverter.ToInt16(fileBytes, i + FECHAPREP_OFFSET);
                    short fechaProRaw  = BitConverter.ToInt16(fileBytes, i + FECHAPRO_OFFSET);
                    DateTime? fechaPrep = ConvertFechatexToDateTime(fechaPrepRaw);
                    DateTime? fechaPro  = ConvertFechatexToDateTime(fechaProRaw);

                    var entity = new Instruci
                    {
                        CodiEmpr    = 0,
                        Codigo      = string.IsNullOrEmpty(codigo)      ? null : codigo,
                        Descripcion = string.IsNullOrEmpty(descripcion) ? null : descripcion,
                        Revision    = string.IsNullOrEmpty(revision)    ? null : revision,
                        Cod_Inte    = codInte,
                        Preparo     = string.IsNullOrEmpty(preparo)     ? null : preparo,
                        FechaPrep   = fechaPrep,
                        Aprobo      = string.IsNullOrEmpty(aprobo)      ? null : aprobo,
                        FechaPro    = fechaPro,
                        Observaciones = null
                        // AddRecord, LastUpdate, CodiEmprNet, NumUsuar → asignados por el AuditInterceptor
                    };

                    entities.Add(entity);
                    processedCodes.Add(codInte);
                    count++;
                }

                if (entities.Any())
                {
                    await _context.Instrucis.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Migración INSTRUCI completada: {count} registros migrados, {duplicatesSkipped} duplicados omitidos, {emptySkipped} registros vacíos omitidos.");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Instruci");
                throw;
            }
        }

        /// <summary>
        /// Actualiza el campo Observaciones de la tabla INSTRUCI con el contenido
        /// del archivo plano INSTRUCD.DAT (cuerpo de texto de cada instrucción).
        ///
        /// DEBE ejecutarse DESPUÉS de MigrateInstruciAsync.
        ///
        /// ESTRUCTURA DEL ARCHIVO PLANO (128 bytes por registro):
        /// - Offset  0 ( 2 bytes): Cod_Inte → Int16 CVI (clave hacia INSTRUCI.Cod_Inte)
        /// - Offset  2 (126 bytes): Linea    → char(126) texto de la línea
        ///
        /// Cada instrucción tiene N líneas. Se concatenan con salto de línea
        /// y se almacenan en INSTRUCI.Observaciones (nvarchar MAX).
        /// </summary>
        public async Task<int> MigrateInstrucdAsync(string? basePath = null)
        {
            try
            {
                // Guardia: si ya hay registros con Observaciones en INSTRUCI, saltar
                var conObservaciones = await _context.Instrucis
                    .Where(x => !string.IsNullOrEmpty(x.Observaciones))
                    .CountAsync();

                if (conObservaciones > 0)
                {
                    _logger.LogInformation($"Ya existen {conObservaciones} registros con Observaciones en INSTRUCI. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "INSTRUCD.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // 128 bytes por registro: 2 bytes Cod_Inte (Int16) + 126 bytes texto de línea
                const int RECORD_LENGTH  = 128;
                const int COD_INTE_OFFSET = 0;
                const int LINEA_OFFSET   = 2;
                const int LINEA_LENGTH   = 126;

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                var enc       = System.Text.Encoding.GetEncoding(1252);

                // Paso 1: agrupar líneas por Cod_Inte manteniendo el orden físico del archivo
                var lineasPorCodInte = new Dictionary<int, List<string>>();

                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    if (fileBytes.Length - i < RECORD_LENGTH) break;

                    short codInteRaw = BitConverter.ToInt16(fileBytes, i + COD_INTE_OFFSET);
                    int codInte = (int)codInteRaw;
                    if (codInte == 0) continue;

                    // TrimEnd solo elimina el relleno de espacios del campo fijo (126 bytes).
                    // Las líneas completamente vacías quedan como "" y se preservan como saltos de párrafo.
                    string linea = enc.GetString(fileBytes, i + LINEA_OFFSET, LINEA_LENGTH).TrimEnd();

                    if (!lineasPorCodInte.ContainsKey(codInte))
                        lineasPorCodInte[codInte] = new List<string>();

                    lineasPorCodInte[codInte].Add(linea);
                }

                // Paso 2: por cada Cod_Inte, buscar el registro en INSTRUCI y actualizar Observaciones
                int count        = 0;
                int notFound     = 0;

                foreach (var kvp in lineasPorCodInte)
                {
                    int codInte = kvp.Key;
                    // Join con Environment.NewLine: las líneas vacías ("") producen líneas en blanco
                    // en el texto final, preservando la estructura de párrafos del documento original.
                    // NO se aplica TrimEnd() al texto final para conservar también los saltos al cierre.
                    string texto = string.Join(Environment.NewLine, kvp.Value);

                    var instruciRec = await _context.Instrucis
                        .FirstOrDefaultAsync(x => x.Cod_Inte == codInte);

                    if (instruciRec != null)
                    {
                        instruciRec.Observaciones = string.IsNullOrWhiteSpace(texto) ? null : texto;
                        count++;
                    }
                    else
                    {
                        notFound++;
                        _logger.LogWarning($"Cod_Inte={codInte} de INSTRUCD no encontrado en tabla INSTRUCI.");
                    }
                }

                if (count > 0)
                    await _context.SaveChangesAsync();

                _logger.LogInformation($"INSTRUCD completado: {count} registros de INSTRUCI actualizados con Observaciones. No encontrados: {notFound}.");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la actualización de Observaciones desde INSTRUCD");
                throw;
            }
        }

        /// <summary>
        /// Convierte una fecha en formato VB6 FECHATEX$ a DateTime?
        /// 
        /// La función FECHATEX$ de VB6 usa una codificación CUSTOM (NO es OLE Automation date):
        ///   FEX = Día + 32 * Mes + 416 * (Año2d - 80)
        /// 
        /// Decodificación (equivalente al algoritmo VB6):
        ///   year2d  = Int(FEX / 416) + 80    → año de 2 dígitos (base 80)
        ///   while year2d > 99: year2d -= 100  → ajuste año 2000
        ///   remaining = FEX - 416 * (year2d - 80)
        ///   month   = remaining / 32
        ///   day     = remaining mod 32
        ///   Siglo: year2d >= 80 → 1900 + year2d ; year2d &lt; 80 → 2000 + year2d
        /// 
        /// Ejemplos verificados:
        ///   FEX = 6447 → Día=15, Mes=6 (Jun), Año=1995
        ///   FEX = 17925 → Día=5,  Mes=1 (Ene), Año=2023
        /// 
        /// Si FEX &lt;= 0 retorna null (fecha no asignada)
        /// </summary>
        private static DateTime? ConvertFechatexToDateTime(short fex)
        {
            if (fex <= 0) return null;
            try
            {
                int fexVal = (int)fex;

                // Extraer año de 2 dígitos (base 80)
                int year2d = fexVal / 416 + 80;

                // Quitar componente año del valor
                int remaining = fexVal - 416 * (year2d - 80);

                // Equivalente al While FEXX(3) > 99: FEXX(3) = FEXX(3) - 100
                while (year2d > 99) year2d -= 100;

                // Extraer mes y día
                int month = remaining / 32;
                int day   = remaining - 32 * month;

                // Determinar siglo: >= 80 → 1900s, < 80 → 2000s
                int year = year2d >= 80 ? 1900 + year2d : 2000 + year2d;

                // Validar componentes antes de crear la fecha
                if (day < 1 || day > 31 || month < 1 || month > 12) return null;

                return new DateTime(year, month, day);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Actualiza el campo Observaciones de la tabla INSTRUME desde el archivo OBINSTRU.DAT.
        ///
        /// DEBE ejecutarse DESPUÉS de MigrateInstrumeAsync.
        ///
        /// ESTRUCTURA DEL ARCHIVO PLANO (128 bytes por registro):
        /// - Offset  0 ( 2 bytes): CVI (Int16) → clave que coincide con Instrume.CodIntMedIdent
        /// - Offset  2 (126 bytes): Línea de texto (Windows-1252, padding de espacios al final)
        ///
        /// Cada instrumento puede tener N líneas. Se concatenan con Environment.NewLine
        /// preservando las líneas vacías que representan saltos de párrafo/página.
        /// </summary>
        public async Task<int> MigrateObInstruAsync(string? basePath = null)
        {
            try
            {
                // Guardia: si ya hay instrumentos con Observaciones, saltar
                var conObservaciones = await _context.Instrumes
                    .Where(x => !string.IsNullOrEmpty(x.Observaciones))
                    .CountAsync();

                if (conObservaciones > 0)
                {
                    _logger.LogInformation($"Ya existen {conObservaciones} registros con Observaciones en INSTRUME. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "OBINSTRU.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // 128 bytes por registro: 2 bytes CVI (CodIntMedIdent) + 126 bytes texto de línea
                const int RECORD_LENGTH  = 128;
                const int KEY_OFFSET     = 0;
                const int LINEA_OFFSET   = 2;
                const int LINEA_LENGTH   = 126;

                var fileBytes = await File.ReadAllBytesAsync(filePath);
                var enc       = System.Text.Encoding.GetEncoding(1252);

                // Paso 1: agrupar líneas por CodIntMedIdent manteniendo el orden físico del archivo
                var lineasPorKey = new Dictionary<int, List<string>>();

                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    if (fileBytes.Length - i < RECORD_LENGTH) break;

                    short keyRaw = BitConverter.ToInt16(fileBytes, i + KEY_OFFSET);
                    int   key    = (int)keyRaw;
                    if (key == 0) continue;

                    // TrimEnd sin argumentos elimina el relleno de espacios del campo fijo.
                    // Las líneas completamente vacías quedan como "" y preservan saltos de párrafo.
                    string linea = enc.GetString(fileBytes, i + LINEA_OFFSET, LINEA_LENGTH).TrimEnd();

                    if (!lineasPorKey.ContainsKey(key))
                        lineasPorKey[key] = new List<string>();

                    lineasPorKey[key].Add(linea);
                }

                // Paso 2: buscar en INSTRUME por CodIntMedIdent y actualizar Observaciones
                int count    = 0;
                int notFound = 0;

                foreach (var kvp in lineasPorKey)
                {
                    int    keyVal = kvp.Key;
                    string texto  = string.Join(Environment.NewLine, kvp.Value);

                    var instrumes = await _context.Instrumes
                        .Where(x => x.CodIntMedIdent == keyVal)
                        .ToListAsync();

                    if (instrumes.Any())
                    {
                        foreach (var rec in instrumes)
                        {
                            rec.Observaciones = string.IsNullOrWhiteSpace(texto) ? null : texto;
                            count++;
                        }
                    }
                    else
                    {
                        notFound++;
                        _logger.LogWarning($"CodIntMedIdent={keyVal} de OBINSTRU no encontrado en tabla INSTRUME.");
                    }
                }

                if (count > 0)
                    await _context.SaveChangesAsync();

                _logger.LogInformation($"OBINSTRU completado: {count} registro(s) de INSTRUME actualizados con Observaciones. No encontrados: {notFound}.");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la actualización de Observaciones desde OBINSTRU");
                throw;
            }
        }

        public async Task<int> MigrateInstrumeAsync(string? basePath = null)
        {
            try
            {
                if (await _context.Instrumes.AnyAsync())
                {
                    _logger.LogInformation("La tabla Instrume no está vacía. Se omite la migración.");
                    return 0;
                }

                string path = basePath
                    ?? _configuration["LegacyDataPath:BasePath"]
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "INSTRUME.DAT");

                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // INSTRUME.DAT: registros de longitud fija organizados en PARES de 256 bytes cada uno.
                //
                // Registro PRINCIPAL (offset par, CVI != 0) — 256 bytes:
                //   Pos   1, len 12 → CODIGO          (string)
                //   Pos  13, len 52 → DESCRIPCION     (string)
                //   Pos  65, len  2 → CodIntMedIdent  (Int16 - CVI)
                //   Pos  67, len  2 → NROGRUPO        (Int16 - CVI)
                //   Pos  69, len  2 → CodIntMedCalib  (Int16 - CVI)
                //   Pos  71, len  2 → (relleno)
                //   Pos  73, len 24 → TIPO
                //   Pos  97, len 24 → MARCA
                //   Pos 121, len 24 → MODELO
                //   Pos 145, len 16 → NRO DE SERIE
                //   Pos 161, len 32 → RANGO
                //   Pos 193, len 32 → POSICION
                //   Pos 225, len 32 → UBICACION
                //
                // Registro EXTENSIÓN (offset par+256, CVI == 0) — 256 bytes:
                //   Bytes 0-47 (48 bytes) → INTERVCALIBRA (texto libre, puede desbordarse desde CODIGO al DESCRIPCION)
                const int RECORD_LENGTH        = 256;  // cada bloque individual
                const int PAIR_LENGTH          = 512;  // par = registro principal + extensión
                const int CODIGO_START         = 0;
                const int CODIGO_LEN           = 12;
                const int DESCRIPCION_START    = 12;
                const int DESCRIPCION_LEN      = 52;
                const int CODINTMEDIDENT_START = 64;
                const int NROGRUPO_START       = 66;
                const int CODINTMEDCALIB_START = 68;
                const int TIPO_START           = 72;
                const int TIPO_LEN             = 24;
                const int MARCA_START          = 96;
                const int MARCA_LEN            = 24;
                const int MODELO_START         = 120;
                const int MODELO_LEN           = 24;
                const int NROSERIE_START       = 144;
                const int NROSERIE_LEN         = 16;
                const int RANGO_START          = 160;
                const int RANGO_LEN            = 32;
                const int POSICION_START       = 192;
                const int POSICION_LEN         = 32;
                const int UBICACION_START      = 224;
                const int UBICACION_LEN        = 32;
                const int INTERVCALIBRA_LEN    = 48;   // primeros 48 bytes del registro extensión

                var enc = System.Text.Encoding.GetEncoding(1252);
                var fileBytes = await File.ReadAllBytesAsync(filePath);
                int count = 0;
                int duplicatesSkipped = 0;
                var entities = new List<Instrume>();
                var processedCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 0; i + PAIR_LENGTH <= fileBytes.Length; i += PAIR_LENGTH)
                {
                    // --- Registro principal ---
                    string codigo = enc.GetString(fileBytes, i + CODIGO_START, CODIGO_LEN).Trim();
                    if (string.IsNullOrEmpty(codigo)) continue;

                    if (processedCodes.Contains(codigo))
                    {
                        duplicatesSkipped++;
                        _logger.LogWarning($"Código duplicado omitido: {codigo}");
                        continue;
                    }

                    string descripcion    = enc.GetString(fileBytes, i + DESCRIPCION_START,   DESCRIPCION_LEN).Trim();
                    short  codIntMedIdent = BitConverter.ToInt16(fileBytes, i + CODINTMEDIDENT_START);
                    short  nroGrupo       = BitConverter.ToInt16(fileBytes, i + NROGRUPO_START);
                    short  codIntMedCalib = BitConverter.ToInt16(fileBytes, i + CODINTMEDCALIB_START);
                    string tipo           = enc.GetString(fileBytes, i + TIPO_START,    TIPO_LEN).Trim();
                    string marca          = enc.GetString(fileBytes, i + MARCA_START,   MARCA_LEN).Trim();
                    string modelo         = enc.GetString(fileBytes, i + MODELO_START,  MODELO_LEN).Trim();
                    string nroSerie       = enc.GetString(fileBytes, i + NROSERIE_START, NROSERIE_LEN).Trim();
                    string rango          = enc.GetString(fileBytes, i + RANGO_START,   RANGO_LEN).Trim();
                    string posicion       = enc.GetString(fileBytes, i + POSICION_START, POSICION_LEN).Trim();
                    string ubicacion      = enc.GetString(fileBytes, i + UBICACION_START, UBICACION_LEN).Trim();

                    // --- Registro extensión (offset +256): contiene INTERVCALIBRA en sus primeros 48 bytes ---
                    int extOffset = i + RECORD_LENGTH;
                    string intervCalibra  = enc.GetString(fileBytes, extOffset, INTERVCALIBRA_LEN).TrimEnd();

                    var entity = new Instrume
                    {
                        Codigo          = codigo,
                        Descripcion     = string.IsNullOrEmpty(descripcion)   ? null : descripcion,
                        CodIntMedIdent  = codIntMedIdent == 0 ? null : (int?)codIntMedIdent,
                        NroGrupo        = nroGrupo       == 0 ? null : (int?)nroGrupo,
                        CodIntMedCalib  = codIntMedCalib  == 0 ? null : (int?)codIntMedCalib,
                        Tipo            = string.IsNullOrEmpty(tipo)          ? null : tipo,
                        Marca           = string.IsNullOrEmpty(marca)         ? null : marca,
                        Modelo          = string.IsNullOrEmpty(modelo)        ? null : modelo,
                        NroSerie        = string.IsNullOrEmpty(nroSerie)      ? null : nroSerie,
                        Rango           = string.IsNullOrEmpty(rango)         ? null : rango,
                        Posicion        = string.IsNullOrEmpty(posicion)      ? null : posicion,
                        Ubicacion       = string.IsNullOrEmpty(ubicacion)     ? null : ubicacion,
                        IntervCalibra   = string.IsNullOrEmpty(intervCalibra) ? null : intervCalibra,
                        Status          = 0
                    };

                    entities.Add(entity);
                    processedCodes.Add(codigo);
                    count++;
                }

                if (entities.Any())
                {
                    await _context.Instrumes.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Se migraron {count} registros a la tabla Instrume. Duplicados omitidos: {duplicatesSkipped}");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Instrume");
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

        /// <summary>
        /// Migración de datos desde BRECEPEN.DAT a la tabla BRECEPEN
        /// Equivalente al código VB6 GENERES_ANUCIEOF
        /// 
        /// Lógica VB6:
        /// - FIN& = ULTREG("BRECEPEN") 'retorna el ultimo registro
        /// - For U& = 1 To FIN&
        /// - TXTX$ = REGLEIDO$("BRECEPEN", U&)
        /// - CANTIENT1 = CVD(Mid$(TXTX$, 88, 8))
        /// - CANTISAL1 = CVD(Mid$(TXTX$, 96, 8))
        /// - CANTI = CANTIENT1 - CANTISAL1
        /// - IF CANTI >0.05 THEN
        /// - FF% = CVI(Mid$(TXTX$, 1, 2)) 'fecha (se usa CVDAT)
        /// - CIII% = CVI(Mid$(TXTX$, 3, 2)) 'COD_INTE
        /// - LTX$ = Trim$(Mid$(TXTX$, 5, 16)) 'IDENLOTE
        /// - CMOV$ = Trim$(Mid$(TXTX$, 21, 2)) 'CODIMOVI
        /// - DOPRICO$ = Trim$(Mid$(TXTX$, 23, 10)) 'DOPRICOR y DOPRILAR
        /// - DOSECO$ = Trim$(Mid$(TXTX$, 33, 10)) 'DOSECCOR y DOSECLAR
        /// - REFE$ = Trim$(Mid$(TXTX$, 43, 30)) 'REFERCOR
        /// - DEPO% = Asc(Mid$(TXTX$, 83, 1)) 'DEPOMOVI
        /// - NC% = CVI(Mid$(TXTX$, 86, 2)) 'NUME_PROV (default 0)
        /// - MONEVAL = 1
        /// - NUME_CLIE = 0
        /// - CODIEMPNET = 101
        /// - NUMUSUAR = 10
        /// - END IF
        /// - Next U&
        /// </summary>
        public async Task<int> MigrateBrecepenAsync(string? basePath = null)
        {
            try
            {
                // Validar que la tabla esté vacía
                if (await _context.Brecepens.AnyAsync())
                {
                    _logger.LogInformation("La tabla Brecepen no está vacía. Se omite la migración.");
                    return 0;
                }

                string path = basePath 
                    ?? _configuration["LegacyDataPath:BasePath"] 
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "BRECEPEN.DAT");
                
                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // El archivo DAT tiene registros de longitud fija sin saltos de línea
                // Longitud total del registro: 128 bytes
                // (BRECEPEN copia la estructura de MOVISTO. El registro se arma en VB6 con
                //  campos hasta la posición 127-128, por lo que la longitud real es 128)
                const int RECORD_LENGTH = 128;
                
                var fileBytes = await File.ReadAllBytesAsync(filePath);
                int count = 0;
                var entities = new List<Brecepen>();

                // Procesar el archivo en bloques de tamaño fijo
                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    // Verificar que hay suficientes bytes para un registro completo
                    int remainingBytes = fileBytes.Length - i;
                    if (remainingBytes < RECORD_LENGTH) break;

                    // Usar codificación Windows-1252 (Latin-1) para caracteres acentuados del español
                    string record = System.Text.Encoding.GetEncoding(1252).GetString(fileBytes, i, RECORD_LENGTH);

                    // Extraer campos según el código VB6
                    // FECHA% = CVI(Mid$(TXTX$, 1, 2)) - fecha empaquetada (se decodifica con FECHATEX)
                    short ffValue = BitConverter.ToInt16(fileBytes, i); // CVI de pos 1-2
                    DateTime? fechaMovimiento = ConvertFechaTex(ffValue);

                    // CIII% = CVI(Mid$(TXTX$, 3, 2)) - COD_INTE
                    short codInte = BitConverter.ToInt16(fileBytes, i + 2); // CVI de pos 3-4

                    // LTX$ = Trim$(Mid$(TXTX$, 5, 16)) - IDENLOTE
                    string idenLote = record.Substring(4, 16).Trim();

                    // CMOV$ = Trim$(Mid$(TXTX$, 21, 2)) - CODIMOVI
                    string codiMovi = record.Substring(20, 2).Trim();

                    // DOPRICO$ = Trim$(Mid$(TXTX$, 23, 10)) - DOPRICOR y DOPRILAR
                    string dopriCor = record.Substring(22, 10).Trim();

                    // DOSECO$ = Trim$(Mid$(TXTX$, 33, 10)) - DOSECCOR y DOSECLAR
                    string doseCor = record.Substring(32, 10).Trim();

                    // REFE$ = Trim$(Mid$(TXTX$, 43, 30)) - REFERCOR
                    string referCor = record.Substring(42, 30).Trim();

                    // DEPO% = Asc(Mid$(TXTX$, 83, 1)) - DEPOMOVI
                    short depoMovi = (short)fileBytes[i + 82]; // Asc de pos 83

                    // NC% = CVI(Mid$(TXTX$, 86, 2)) - NUME_PROV (default 0)
                    short numeProv = BitConverter.ToInt16(fileBytes, i + 85); // CVI de pos 86-87

                    // CANTIENT1 = CVD(Mid$(TXTX$, 88, 8)) - CANTINGRE (MKD$ = double 8 bytes MBF)
                    double cantIngre = ConvertMBFDoubleToDouble(fileBytes.Skip(i + 87).Take(8).ToArray());

                    // CANTISAL1 = CVD(Mid$(TXTX$, 96, 8)) - CANTSALID (MKD$ = double 8 bytes MBF)
                    double cantSalid = ConvertMBFDoubleToDouble(fileBytes.Skip(i + 95).Take(8).ToArray());

                    // CANTI = CANTIENT1 - CANTISALID
                    // Validar que los valores estén en rango para decimal
                    decimal canti = 0;
                    try
                    {
                        decimal cantIngreDec = (decimal)cantIngre;
                        decimal cantSalidDec = (decimal)cantSalid;
                        canti = cantIngreDec - cantSalidDec;
                    }
                    catch (OverflowException)
                    {
                        // Si el valor está fuera de rango, omitir este registro
                        _logger.LogWarning($"Registro omitido: valor fuera de rango para decimal. CantIngre: {cantIngre}, CantSalid: {cantSalid}");
                        continue;
                    }

                    // IF CANTI >0.05 THEN
                    if (canti > 0.05m)
                    {
                        // Validar conversión a decimal para Cantingre y Cantsalid
                        decimal? cantIngreDec = null;
                        decimal? cantSalidDec = null;
                        
                        try
                        {
                            cantIngreDec = (decimal)cantIngre;
                        }
                        catch (OverflowException)
                        {
                            _logger.LogWarning($"CantIngre fuera de rango: {cantIngre}, se usará null");
                        }
                        
                        try
                        {
                            cantSalidDec = (decimal)cantSalid;
                        }
                        catch (OverflowException)
                        {
                            _logger.LogWarning($"CantSalid fuera de rango: {cantSalid}, se usará null");
                        }

                        var entity = new Brecepen
                        {
                            Fechmov = fechaMovimiento,
                            Feremovi = fechaMovimiento, // FEREMOVI usa la misma fecha
                            COD_INTE = codInte,
                            Idenlote = idenLote,
                            Codimovi = codiMovi,
                            Doprilar = dopriCor,      // DOPRILAR
                            Doseclar = doseCor,        // DOSECLAR
                            Dopricor = dopriCor,      // DOPRICOR
                            Doseccor = doseCor,        // DOSECCOR
                            Refercor = referCor,
                            Depomovi = depoMovi,
                            NumeProv = numeProv > 0 ? numeProv : 0,
                            Cantingre = cantIngreDec,
                            Cantsalid = cantSalidDec,
                            Moneval = 1,              // MONEVAL = 1
                            NumeClie = 0,              // NUME_CLIE = 0
                            CodiEmprNet = 101,         // CODIEMPNET = 101
                            NumUsuar = 10              // NUMUSUAR = 10
                        };

                        entities.Add(entity);
                        count++;
                    }
                }

                if (entities.Any())
                {
                    await _context.Brecepens.AddRangeAsync(entities);
                    await _context.SaveChangesAsync();
                }

                _logger.LogInformation($"Se migraron {count} registros a la tabla Brecepen.");
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la migración de Brecepen");
                throw;
            }
        }

        /// <summary>
        /// Convierte un número de fecha VB6 (días desde 30/12/1899) a DateTime
        /// Equivalente a la función CVDAT de VB6
        /// </summary>
        private DateTime ConvertVB6Date(short vb6Date)
        {
            // VB6 usa fecha base 30/12/1899
            DateTime baseDate = new DateTime(1899, 12, 30);
            return baseDate.AddDays(vb6Date);
        }

        /// <summary>
        /// Decodifica una fecha empaquetada de VB6 (FLEXOFT) equivalente a la función FECHATEX$.
        /// El entero (leído con CVI) codifica: FEX = dia + 32*mes + 416*(anio-80)
        /// - anio = Int(FEX/416) + 80  (con ajuste año 2000: mientras > 99 se resta 100)
        /// - resto = FEX - 416*(anio-80)
        /// - mes = Int(resto/32)
        /// - dia = resto - 32*mes
        /// El siglo es 1900 si el año de 2 dígitos es >= 80, de lo contrario 2000.
        /// Si la fecha resultante es inválida, VB6 decrementa FEX en 1 y reintenta.
        /// </summary>
        private DateTime? ConvertFechaTex(int fex)
        {
            if (fex <= 0) return null;

            int intentos = 0;
            int fexActual = fex;

            while (fexActual > 0 && intentos < 60)
            {
                int anio = (fexActual / 416) + 80;
                int resto = fexActual - 416 * (anio - 80);
                int mes = resto / 32;
                int dia = resto - 32 * mes;

                // Ajuste año 2000: reducir a 2 dígitos
                while (anio > 99) anio -= 100;

                // Determinar el siglo: >=80 -> 1900, si no -> 2000
                int siglo = anio >= 80 ? 1900 : 2000;
                int anioCompleto = siglo + anio;

                if (mes >= 1 && mes <= 12 && dia >= 1 && dia <= 31)
                {
                    try
                    {
                        return new DateTime(anioCompleto, mes, dia);
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        // Fecha inválida (ej. 31 de un mes de 30 días): reintentar con FEX-1
                    }
                }

                fexActual--;
                intentos++;
            }

            return null;
        }

        /// <summary>
        /// Convierte un número en formato Microsoft Binary Format (MBF) de 8 bytes (double)
        /// a un double IEEE. Equivalente a la función CVD de VB6 (MKD$ escribe este formato).
        /// 
        /// Estructura MBF double (8 bytes):
        /// - byte[7]: exponente sesgado (bias 128). Si es 0, el valor es 0.
        /// - bit 7 de byte[6]: signo (1 = negativo)
        /// - mantisa de 55 bits distribuida en byte[6] (7 bits) .. byte[0]
        /// </summary>
        private double ConvertMBFDoubleToDouble(byte[] mbfBytes)
        {
            if (mbfBytes == null || mbfBytes.Length != 8)
                return 0d;

            byte exponente = mbfBytes[7];

            // Exponente 0 -> valor cero
            if (exponente == 0)
                return 0d;

            // Signo: bit más significativo de byte[6]
            int signo = (mbfBytes[6] & 0x80) == 0 ? 1 : -1;

            // Exponente real (bias 128 en MBF; el 1.0 implícito se agrega abajo)
            int exp = exponente - 129;

            // Construir la mantisa (55 bits) sin el 1.0 implícito.
            // byte[6] aporta 7 bits (se quita el bit de signo), luego byte[5]..byte[0]
            double mantisa = (mbfBytes[6] & 0x7F) / 128.0;
            mantisa += mbfBytes[5] / 32768.0;
            mantisa += mbfBytes[4] / 8388608.0;
            mantisa += mbfBytes[3] / 2147483648.0;
            mantisa += mbfBytes[2] / 549755813888.0;
            mantisa += mbfBytes[1] / 140737488355328.0;
            mantisa += mbfBytes[0] / 36028797018963968.0;

            // Resultado = signo * (1 + mantisa) * 2^exp
            return signo * (1.0 + mantisa) * Math.Pow(2, exp);
        }

        /// <summary>
        /// Previsualiza los datos de BRECEPEN.DAT sin migrar
        /// Devuelve una lista de objetos con los datos a migrar
        /// </summary>
        public async Task<List<BrecepenPreviewDTO>> PreviewBrecepenAsync(string? basePath = null)
        {
            try
            {
                string path = basePath 
                    ?? _configuration["LegacyDataPath:BasePath"] 
                    ?? @"C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat";
                string filePath = Path.Combine(path, "BRECEPEN.DAT");
                
                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Archivo no encontrado: {filePath}");
                    throw new FileNotFoundException($"No se encontró el archivo de datos: {filePath}");
                }

                // Longitud real del registro: 128 bytes (BRECEPEN copia estructura de MOVISTO)
                const int RECORD_LENGTH = 128;
                var fileBytes = await File.ReadAllBytesAsync(filePath);
                var data = new List<BrecepenPreviewDTO>();

                // Procesar el archivo en bloques de tamaño fijo
                for (int i = 0; i < fileBytes.Length; i += RECORD_LENGTH)
                {
                    int remainingBytes = fileBytes.Length - i;
                    if (remainingBytes < RECORD_LENGTH) break;

                    string record = System.Text.Encoding.GetEncoding(1252).GetString(fileBytes, i, RECORD_LENGTH);

                    // FECHA% = CVI(pos 1-2) -> se decodifica con FECHATEX
                    short ffValue = BitConverter.ToInt16(fileBytes, i);
                    DateTime? fechaMovimiento = ConvertFechaTex(ffValue);

                    short codInte = BitConverter.ToInt16(fileBytes, i + 2);
                    string idenLote = record.Substring(4, 16).Trim();
                    string codiMovi = record.Substring(20, 2).Trim();
                    string dopriCor = record.Substring(22, 10).Trim();
                    string doseCor = record.Substring(32, 10).Trim();
                    string referCor = record.Substring(42, 30).Trim();
                    short depoMovi = (short)fileBytes[i + 82];
                    short numeProv = BitConverter.ToInt16(fileBytes, i + 85);

                    // CANTINGRE/CANTSALID = MKD$ = double 8 bytes MBF (CVD)
                    double cantIngre = ConvertMBFDoubleToDouble(fileBytes.Skip(i + 87).Take(8).ToArray());
                    double cantSalid = ConvertMBFDoubleToDouble(fileBytes.Skip(i + 95).Take(8).ToArray());

                    decimal canti = 0;
                    try
                    {
                        decimal cantIngreDec = (decimal)cantIngre;
                        decimal cantSalidDec = (decimal)cantSalid;
                        canti = cantIngreDec - cantSalidDec;
                    }
                    catch (OverflowException)
                    {
                        canti = 0;
                    }

                    // Solo mostrar los registros que se van a migrar (saldo > 0.05)
                    if (canti <= 0.05m) continue;

                    data.Add(new BrecepenPreviewDTO
                    {
                        Fecha = fechaMovimiento?.ToString("dd/MM/yyyy") ?? "",
                        CodInte = codInte,
                        IdLote = idenLote,
                        CodMovi = codiMovi,
                        DopriCor = dopriCor,
                        DoseCor = doseCor,
                        ReferCor = referCor,
                        DepoMovi = depoMovi,
                        NumeProv = numeProv,
                        CantIngre = cantIngre,
                        CantSalid = cantSalid,
                        Saldo = canti,
                        Migrar = true
                    });
                }

                _logger.LogInformation($"Se previsualizaron {data.Count} registros de BRECEPEN.DAT");
                return data;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error durante la previsualización de Brecepen");
                throw;
            }
        }
    }
}
