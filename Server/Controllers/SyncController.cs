using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para sincronización entre Central y Sucursal
    /// </summary>
    [ApiController]
    [Route("api/sync")]
    public class SyncController : ControllerBase
    {
        private readonly ISyncService _syncService;
        private readonly ILogger<SyncController> _logger;
        private readonly DataContext _context;

        public SyncController(ISyncService syncService, ILogger<SyncController> logger, DataContext context)
        {
            _syncService = syncService;
            _logger = logger;
            _context = context;
        }

        /// <summary>
        /// Obtiene la configuración actual de sincronización
        /// </summary>
        [HttpGet("config")]
        // [Authorize(Roles = "Admin")]
        [Authorize(Roles = "Administrador,SUPERADMIN")]
        public ActionResult<SyncConfigDTO> GetConfig()
        {
            try
            {
                var config = _syncService.ObtenerConfiguracion();
                return Ok(config);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener configuración de sincronización");
                return StatusCode(500, $"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Guarda la configuración de sincronización
        /// </summary>
        [HttpPost("config")]
        //[Authorize(Roles = "Admin")]
        [Authorize(Roles = "Administrador,SUPERADMIN")]
        public async Task<IActionResult> SaveConfig([FromBody] SyncConfigDTO config)
        {
            try
            {
                var resultado = await _syncService.GuardarConfiguracionAsync(config);
                if (resultado)
                {
                    return Ok(new { mensaje = "Configuración guardada exitosamente" });
                }
                else
                {
                    return BadRequest("No se pudo guardar la configuración");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al guardar configuración de sincronización");
                return StatusCode(500, $"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Procesa archivos de sincronización pendientes (Solo Sucursales)
        /// 
        /// Este endpoint permite a una Sucursal procesar manualmente los archivos JSON
        /// generados por la Central en la carpeta de intercambio compartida.
        /// 
        /// Proceso:
        /// - Busca archivos MASTER_*.json, FORMULAS_*.json, SECOPER_*.json
        /// - Verifica qué archivos ya fueron procesados (tabla SyncProcesados)
        /// - Procesa archivos pendientes en orden cronológico
        /// - Actualiza tablas Master, Formulas y SecOper
        /// 
        /// Nota: No requiere autorización específica porque la validación de permisos
        /// se maneja en el frontend (botón solo visible en Sucursales con configuración).
        /// </summary>
        [HttpPost("procesar")]
        public async Task<ActionResult<SyncResultDTO>> ProcesarSync([FromBody] ValidarRutaDTO request)
        {
            try
            {
                var resultado = await _syncService.ProcesarArchivosSyncAsync(request.Ruta);
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar archivos de sincronización");
                return StatusCode(500, new SyncResultDTO
                {
                    Exitoso = false,
                    Mensaje = $"Error: {ex.Message}"
                });
            }
        }

        /// <summary>
        /// Valida que una ruta de intercambio sea accesible
        /// </summary>
        [HttpPost("validar-ruta")]
        // [Authorize(Roles = "Admin")]
        [Authorize(Roles = "Administrador,SUPERADMIN")]
        public ActionResult<ValidarRutaResponseDTO> ValidarRuta([FromBody] ValidarRutaDTO request)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.Ruta))
                {
                    return BadRequest("La ruta no puede estar vacía");
                }

                if (Directory.Exists(request.Ruta))
                {
                    return Ok(new ValidarRutaResponseDTO { Valida = true, Mensaje = "Ruta accesible" });
                }
                else
                {
                    return Ok(new ValidarRutaResponseDTO { Valida = false, Mensaje = "La ruta no existe o no es accesible" });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al validar ruta: {request.Ruta}");
                return Ok(new ValidarRutaResponseDTO { Valida = false, Mensaje = $"Error: {ex.Message}" });
            }
        }

        /// <summary>
        /// Procesa archivos de sincronización de remitos pendientes (Solo Central)
        ///
        /// Este endpoint permite a la Central procesar manualmente los archivos JSON
        /// de remitos generados por la Sucursal de Neuquén en la carpeta de intercambio compartida.
        ///
        /// Proceso:
        /// - Busca archivos REMITO_*.json
        /// - Verifica qué archivos ya fueron procesados (tabla SyncProcesados)
        /// - Procesa archivos pendientes en orden cronológico
        /// - Genera INSERT en la tabla MOVISTO
        ///
        /// Nota: No requiere autorización específica porque la validación de permisos
        /// se maneja en el frontend (botón solo visible en Central con configuración).
        /// </summary>
        [HttpPost("procesar-remitos")]
        public async Task<ActionResult<SyncResultDTO>> ProcesarRemitos([FromBody] ValidarRutaDTO request)
        {
            try
            {
                // Obtener el número de usuario activo del contexto de autenticación
                // NOTA: Esto debe implementarse según el sistema de autenticación del proyecto
                var numUsuar = 0; // Valor por defecto, debe obtenerse del usuario autenticado

                _logger.LogInformation("Iniciando procesamiento manual de remitos en ruta: {Ruta}", request.Ruta);

                // PRIMERO: Procesar archivos TXT y convertirlos a JSON
                if (!string.IsNullOrWhiteSpace(request.Ruta) && Directory.Exists(request.Ruta))
                {
                    var archivosTxt = Directory.GetFiles(request.Ruta, "*.txt");
                    _logger.LogInformation("Archivos TXT encontrados: {Cantidad}", archivosTxt.Length);

                    foreach (var archivoTxt in archivosTxt)
                    {
                        var nombreArchivo = Path.GetFileName(archivoTxt);
                        _logger.LogInformation("Convirtiendo archivo TXT a JSON: {Archivo}", nombreArchivo);

                        try
                        {
                            var resultadoConversion = await _syncService.GenerarArchivoSyncRemitoDesdeTxtAsync(archivoTxt, numUsuar);

                            if (resultadoConversion)
                            {
                                _logger.LogInformation("Archivo TXT convertido exitosamente: {Archivo}", nombreArchivo);
                            }
                            else
                            {
                                _logger.LogWarning("Error al convertir archivo TXT: {Archivo}", nombreArchivo);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Excepción al convertir archivo TXT: {Archivo}", nombreArchivo);
                        }
                    }
                }

                // SEGUNDO: Procesar archivos JSON de remitos
                var resultado = await _syncService.ProcesarArchivosRemitosAsync(request.Ruta, numUsuar);
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar archivos de remitos");
                return StatusCode(500, new SyncResultDTO
                {
                    Exitoso = false,
                    Mensaje = $"Error: {ex.Message}"
                });
            }
        }

        /// <summary>
        /// Procesa archivos de sincronización de órdenes de reparación pendientes (Solo Central)
        ///
        /// Este endpoint permite a la Central procesar manualmente los archivos JSON
        /// de órdenes de reparación generados por la Sucursal en la carpeta de intercambio compartida.
        ///
        /// Proceso:
        /// - Busca archivos REPARACION_*.json y archivos TXT con prefijo OR.
        /// - Convierte archivos TXT a JSON
        /// - Verifica qué archivos ya fueron procesados (tabla SyncProcesados)
        /// - Procesa archivos pendientes en orden cronológico
        /// - Genera INSERT en la tabla MOVISTO
        ///
        /// Nota: No requiere autorización específica porque la validación de permisos
        /// se maneja en el frontend (botón solo visible en Central con configuración).
        /// </summary>
        [HttpPost("procesar-reparaciones")]
        public async Task<ActionResult<SyncResultDTO>> ProcesarReparaciones([FromBody] ValidarRutaDTO request)
        {
            try
            {
                // Obtener el número de usuario activo del contexto de autenticación
                // NOTA: Esto debe implementarse según el sistema de autenticación del proyecto
                var numUsuar = 0; // Valor por defecto, debe obtenerse del usuario autenticado

                _logger.LogInformation("Iniciando procesamiento manual de reparaciones en ruta: {Ruta}", request.Ruta);

                // PRIMERO: Procesar archivos TXT y convertirlos a JSON
                if (!string.IsNullOrWhiteSpace(request.Ruta) && Directory.Exists(request.Ruta))
                {
                    var archivosTxtOrPunto = Directory.GetFiles(request.Ruta, "OR.*.txt");
                    var archivosTxtOrGuion = Directory.GetFiles(request.Ruta, "OR-*.txt");
                    var archivosTxt = archivosTxtOrPunto.Concat(archivosTxtOrGuion).ToArray();
                    _logger.LogInformation("Archivos TXT de reparaciones encontrados: {Cantidad}", archivosTxt.Length);

                    foreach (var archivoTxt in archivosTxt)
                    {
                        var nombreArchivo = Path.GetFileName(archivoTxt);
                        _logger.LogInformation("Convirtiendo archivo TXT de reparación a JSON: {Archivo}", nombreArchivo);

                        try
                        {
                            var resultadoConversion = await _syncService.GenerarArchivoSyncReparacionDesdeTxtAsync(archivoTxt, numUsuar);

                            if (resultadoConversion)
                            {
                                _logger.LogInformation("Archivo TXT de reparación convertido exitosamente: {Archivo}", nombreArchivo);
                            }
                            else
                            {
                                _logger.LogWarning("Error al convertir archivo TXT de reparación: {Archivo}", nombreArchivo);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Excepción al convertir archivo TXT de reparación: {Archivo}", nombreArchivo);
                        }
                    }
                }

                // SEGUNDO: Procesar archivos JSON de reparaciones
                var resultado = await _syncService.ProcesarArchivosReparacionesAsync(request.Ruta, numUsuar);
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar archivos de reparaciones");
                return StatusCode(500, new SyncResultDTO
                {
                    Exitoso = false,
                    Mensaje = $"Error: {ex.Message}"
                });
            }
        }

        /// <summary>
        /// Procesa archivos de sincronización de ajustes de stock pendientes (Solo Central)
        ///
        /// Este endpoint permite a la Central procesar manualmente los archivos JSON
        /// de ajustes de stock generados por la Sucursal en la carpeta de intercambio compartida.
        ///
        /// Proceso:
        /// - Busca archivos AJUSTE_*.json y archivos TXT con prefijo AJ.
        /// - Convierte archivos TXT a JSON
        /// - Verifica qué archivos ya fueron procesados (tabla SyncProcesados)
        /// - Procesa archivos pendientes en orden cronológico
        /// - Genera INSERT en la tabla MOVISTO
        ///
        /// Nota: No requiere autorización específica porque la validación de permisos
        /// se maneja en el frontend (botón solo visible en Central con configuración).
        /// </summary>
        [HttpPost("procesar-ajustes")]
        public async Task<ActionResult<SyncResultDTO>> ProcesarAjustes([FromBody] ValidarRutaDTO request)
        {
            try
            {
                // Obtener el número de usuario activo del contexto de autenticación
                // NOTA: Esto debe implementarse según el sistema de autenticación del proyecto
                var numUsuar = 0; // Valor por defecto, debe obtenerse del usuario autenticado

                _logger.LogInformation("Iniciando procesamiento manual de ajustes en ruta: {Ruta}", request.Ruta);

                // PRIMERO: Procesar archivos TXT y convertirlos a JSON
                if (!string.IsNullOrWhiteSpace(request.Ruta) && Directory.Exists(request.Ruta))
                {
                    var archivosTxtAjPunto = Directory.GetFiles(request.Ruta, "AJ.*.txt");
                    var archivosTxtAjGuion = Directory.GetFiles(request.Ruta, "AJ-*.txt");
                    var archivosTxt = archivosTxtAjPunto.Concat(archivosTxtAjGuion).ToArray();
                    _logger.LogInformation("Archivos TXT de ajustes encontrados: {Cantidad}", archivosTxt.Length);

                    foreach (var archivoTxt in archivosTxt)
                    {
                        var nombreArchivo = Path.GetFileName(archivoTxt);
                        _logger.LogInformation("Convirtiendo archivo TXT de ajuste a JSON: {Archivo}", nombreArchivo);

                        try
                        {
                            var resultadoConversion = await _syncService.GenerarArchivoSyncAjusteDesdeTxtAsync(archivoTxt, numUsuar);

                            if (resultadoConversion)
                            {
                                _logger.LogInformation("Archivo TXT de ajuste convertido exitosamente: {Archivo}", nombreArchivo);
                            }
                            else
                            {
                                _logger.LogWarning("Error al convertir archivo TXT de ajuste: {Archivo}", nombreArchivo);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Excepción al convertir archivo TXT de ajuste: {Archivo}", nombreArchivo);
                        }
                    }
                }

                // SEGUNDO: Procesar archivos JSON de ajustes
                var resultado = await _syncService.ProcesarArchivosAjustesAsync(request.Ruta, numUsuar);
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar archivos de ajustes");
                return StatusCode(500, new SyncResultDTO
                {
                    Exitoso = false,
                    Mensaje = $"Error: {ex.Message}"
                });
            }
        }

        /// <summary>
        /// Procesa archivos de sincronización de pedidos de reposición pendientes (Solo Central)
        ///
        /// Este endpoint permite a la Central procesar manualmente los archivos TXT
        /// de pedidos de reposición generados por la Sucursal en la carpeta de intercambio compartida.
        ///
        /// Proceso:
        /// - Verifica que la configuración de sincronización tenga habilitado ProcesaPedidosReposicion
        /// - Busca archivos Ocompra_Sucursal*.txt
        /// - Verifica qué archivos ya fueron procesados (tabla SyncProcesados)
        /// - Procesa archivos pendientes en orden cronológico
        /// - Inserta o actualiza registros en la tabla OC_SUCURSAL
        ///
        /// Formato del archivo según código VB6:
        /// - Nombre: Ocompra_Sucursal AAAA_MM_DD HHMMSS.TXT
        /// - Delimitador: punto y coma (;)
        /// - Primera línea: encabezado con nombres de columnas
        /// - Líneas siguientes: datos de cada registro
        /// - Última línea opcional: mensaje con formato ||MENSAJE||;texto
        ///
        /// Nota: No requiere autorización específica porque la validación de permisos
        /// se maneja en el frontend (botón solo visible en Central con configuración).
        /// </summary>
        [HttpPost("procesar-pedidos-reposicion")]
        public async Task<ActionResult<SyncResultDTO>> ProcesarPedidosReposicion([FromBody] ValidarRutaDTO request)
        {
            try
            {
                // ====================================================================
                // VALIDAR CONFIGURACIÓN DE SINCRONIZACIÓN
                // ====================================================================
                // Verificar que este equipo esté configurado para procesar pedidos de reposición
                // Esto evita que una sucursal procese archivos que solo debe procesar la central
                var config = _syncService.ObtenerConfiguracion();

                if (!config.ProcesaPedidosReposicion)
                {
                    _logger.LogWarning("Intento de procesar pedidos de reposición sin configuración habilitada");
                    return BadRequest(new SyncResultDTO
                    {
                        Exitoso = false,
                        Mensaje = "Este equipo no está configurado para procesar pedidos de reposición. Por favor habilite la opción ProcesaPedidosReposicion en la configuración de sincronización."
                    });
                }

                // Obtener el número de usuario activo del contexto de autenticación
                // NOTA: Esto debe implementarse según el sistema de autenticación del proyecto
                var numUsuar = 0; // Valor por defecto, debe obtenerse del usuario autenticado

                _logger.LogInformation("Iniciando procesamiento manual de pedidos de reposición en ruta: {Ruta}", request.Ruta);

                // Procesar archivos TXT de pedidos de reposición
                var resultado = await _syncService.ProcesarArchivosOcompraSucursalAsync(request.Ruta, numUsuar);
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar archivos de pedidos de reposición");
                return StatusCode(500, new SyncResultDTO
                {
                    Exitoso = false,
                    Mensaje = $"Error: {ex.Message}"
                });
            }
        }

        /// <summary>
        /// Consulta la cantidad de archivos pendientes de sincronización
        /// Retorna la cantidad de archivos JSON pendientes para productos y remitos
        /// </summary>
        [HttpGet("archivos-pendientes")]
        public ActionResult<SyncArchivosPendientesDTO> GetArchivosPendientes()
        {
            try
            {
                var config = _syncService.ObtenerConfiguracion();
                var resultado = new SyncArchivosPendientesDTO
                {
                    RutaProductos = config.RutaIntercambio,
                    RutaRemitos = config.RutaIntercambioRemitos,
                    RutaReparaciones = config.RutaIntercambioReparaciones,
                    RutaAjustes = config.RutaIntercambioAjustes,
                    ArchivosProductos = 0,
                    ArchivosRemitos = 0,
                    ArchivosReparaciones = 0,
                    ArchivosAjustes = 0
                };

                // Contar archivos de productos pendientes
                if (!string.IsNullOrWhiteSpace(config.RutaIntercambio) && Directory.Exists(config.RutaIntercambio))
                {
                    var archivosProductos = Directory.GetFiles(config.RutaIntercambio, "*.json")
                        .Where(f => Path.GetFileName(f).StartsWith("MASTER_") ||
                                     Path.GetFileName(f).StartsWith("FORMULAS_") ||
                                     Path.GetFileName(f).StartsWith("SECOPER_"))
                        .ToList();

                    // Filtrar archivos ya procesados
                    var archivosYaProcesados = _context.SyncProcesados
                        .Where(s => s.Estado == "EXITOSO" && archivosProductos.Select(Path.GetFileName).Contains(s.Archivo))
                        .Select(s => s.Archivo)
                        .ToList();

                    resultado.ArchivosProductos = archivosProductos.Count(f => !archivosYaProcesados.Contains(Path.GetFileName(f)));
                }

                // Contar archivos de remitos pendientes
                if (!string.IsNullOrWhiteSpace(config.RutaIntercambioRemitos) && Directory.Exists(config.RutaIntercambioRemitos))
                {
                    // Contar archivos JSON pendientes
                    var archivosRemitosJson = Directory.GetFiles(config.RutaIntercambioRemitos, "REMITO_*.json").ToList();
                    _logger.LogInformation("Archivos JSON de remitos encontrados en disco: {Cantidad}", archivosRemitosJson.Count);
                    foreach (var archivo in archivosRemitosJson)
                    {
                        _logger.LogDebug("Archivo JSON: {Archivo}", Path.GetFileName(archivo));
                    }

                    // Obtener todos los archivos procesados de la tabla SyncProcesados
                    var todosProcesados = _context.SyncProcesados
                        .Where(s => s.Estado == "EXITOSO")
                        .Select(s => s.Archivo)
                        .ToList();
                    _logger.LogInformation("Archivos en SyncProcesados con estado EXITOSO: {Cantidad}", todosProcesados.Count);
                    foreach (var proc in todosProcesados)
                    {
                        _logger.LogDebug("Procesado: {Archivo}", proc);
                    }

                    // Filtrar archivos JSON ya procesados
                    var archivosJsonYaProcesados = archivosRemitosJson
                        .Where(f => todosProcesados.Contains(Path.GetFileName(f)))
                        .ToList();
                    _logger.LogInformation("Archivos JSON ya procesados: {Cantidad}", archivosJsonYaProcesados.Count);

                    var archivosJsonPendientes = archivosRemitosJson.Count(f => !todosProcesados.Contains(Path.GetFileName(f)));
                    _logger.LogInformation("Archivos JSON pendientes: {Cantidad}", archivosJsonPendientes);

                    // Contar archivos TXT pendientes (aún no convertidos a JSON)
                    var archivosRemitosTxt = Directory.GetFiles(config.RutaIntercambioRemitos, "*.txt").ToList();
                    _logger.LogInformation("Archivos TXT de remitos encontrados en disco: {Cantidad}", archivosRemitosTxt.Count);

                    // Filtrar archivos TXT ya procesados (convertidos a JSON)
                    var archivosTxtYaProcesados = archivosRemitosTxt
                        .Where(f => todosProcesados.Contains(Path.GetFileName(f)))
                        .ToList();
                    _logger.LogInformation("Archivos TXT ya procesados: {Cantidad}", archivosTxtYaProcesados.Count);

                    var archivosTxtPendientes = archivosRemitosTxt.Count(f => !todosProcesados.Contains(Path.GetFileName(f)));
                    _logger.LogInformation("Archivos TXT pendientes: {Cantidad}", archivosTxtPendientes);

                    resultado.ArchivosRemitos = archivosJsonPendientes + archivosTxtPendientes;

                    _logger.LogInformation("Archivos de remitos pendientes: {JsonPendientes} JSON + {TxtPendientes} TXT = {Total}",
                        archivosJsonPendientes, archivosTxtPendientes, resultado.ArchivosRemitos);
                }

                // Contar archivos de reparaciones pendientes
                if (!string.IsNullOrWhiteSpace(config.RutaIntercambioReparaciones) && Directory.Exists(config.RutaIntercambioReparaciones))
                {
                    // Contar archivos JSON pendientes de reparaciones
                    var archivosReparacionesJson = Directory.GetFiles(config.RutaIntercambioReparaciones, "OR.*.json")
                        .Concat(Directory.GetFiles(config.RutaIntercambioReparaciones, "OR-*.json"))
                        .ToList();
                    _logger.LogInformation("Archivos JSON de reparaciones encontrados en disco: {Cantidad}", archivosReparacionesJson.Count);

                    // Contar archivos TXT pendientes de reparaciones (aún no convertidos a JSON)
                    var archivosReparacionesTxt = Directory.GetFiles(config.RutaIntercambioReparaciones, "OR.*.txt")
                        .Concat(Directory.GetFiles(config.RutaIntercambioReparaciones, "OR-*.txt"))
                        .ToList();
                    _logger.LogInformation("Archivos TXT de reparaciones encontrados en disco: {Cantidad}", archivosReparacionesTxt.Count);

                    // Obtener todos los archivos procesados de la tabla SyncProcesados
                    var todosProcesadosRep = _context.SyncProcesados
                        .Where(s => s.Estado == "EXITOSO")
                        .Select(s => s.Archivo)
                        .ToList();

                    var archivosJsonRepPendientes = archivosReparacionesJson.Count(f => !todosProcesadosRep.Contains(Path.GetFileName(f)));
                    var archivosTxtRepPendientes = archivosReparacionesTxt.Count(f => !todosProcesadosRep.Contains(Path.GetFileName(f)));

                    resultado.ArchivosReparaciones = archivosJsonRepPendientes + archivosTxtRepPendientes;

                    _logger.LogInformation("Archivos de reparaciones pendientes: {JsonPendientes} JSON + {TxtPendientes} TXT = {Total}",
                        archivosJsonRepPendientes, archivosTxtRepPendientes, resultado.ArchivosReparaciones);
                }

                // Contar archivos de ajustes de stock pendientes
                if (!string.IsNullOrWhiteSpace(config.RutaIntercambioAjustes) && Directory.Exists(config.RutaIntercambioAjustes))
                {
                    // Contar archivos JSON pendientes de ajustes
                    var archivosAjustesJson = Directory.GetFiles(config.RutaIntercambioAjustes, "AJ.*.json")
                        .Concat(Directory.GetFiles(config.RutaIntercambioAjustes, "AJ-*.json"))
                        .ToList();
                    _logger.LogInformation("Archivos JSON de ajustes encontrados en disco: {Cantidad}", archivosAjustesJson.Count);

                    // Contar archivos TXT pendientes de ajustes (aún no convertidos a JSON)
                    var archivosAjustesTxt = Directory.GetFiles(config.RutaIntercambioAjustes, "AJ.*.txt")
                        .Concat(Directory.GetFiles(config.RutaIntercambioAjustes, "AJ-*.txt"))
                        .ToList();
                    _logger.LogInformation("Archivos TXT de ajustes encontrados en disco: {Cantidad}", archivosAjustesTxt.Count);

                    // Obtener todos los archivos procesados de la tabla SyncProcesados
                    var todosProcesadosAjustes = _context.SyncProcesados
                        .Where(s => s.Estado == "EXITOSO")
                        .Select(s => s.Archivo)
                        .ToList();

                    var archivosJsonAjustesPendientes = archivosAjustesJson.Count(f => !todosProcesadosAjustes.Contains(Path.GetFileName(f)));
                    var archivosTxtAjustesPendientes = archivosAjustesTxt.Count(f => !todosProcesadosAjustes.Contains(Path.GetFileName(f)));

                    resultado.ArchivosAjustes = archivosJsonAjustesPendientes + archivosTxtAjustesPendientes;

                    _logger.LogInformation("Archivos de ajustes pendientes: {JsonPendientes} JSON + {TxtPendientes} TXT = {Total}",
                        archivosJsonAjustesPendientes, archivosTxtAjustesPendientes, resultado.ArchivosAjustes);
                }

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al consultar archivos pendientes");
                return StatusCode(500, new SyncArchivosPendientesDTO
                {
                    ArchivosProductos = 0,
                    ArchivosRemitos = 0
                });
            }
        }
    }
}
