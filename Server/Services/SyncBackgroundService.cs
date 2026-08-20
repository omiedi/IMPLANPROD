using IMPLANPROD.Server.Services;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio en segundo plano que procesa automáticamente archivos de sincronización
    /// Solo se ejecuta en instalaciones que procesan archivos de productos (Sucursales)
    /// </summary>
    public class SyncBackgroundService : BackgroundService
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly ILogger<SyncBackgroundService> _logger;
        private readonly TimeSpan _intervalo = TimeSpan.FromMinutes(4);

        public SyncBackgroundService(
            IServiceProvider serviceProvider,
            ILogger<SyncBackgroundService> logger)
        {
            _serviceProvider = serviceProvider;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("SyncBackgroundService iniciado");

            // Esperar 30 segundos antes de la primera ejecución para que la app termine de iniciar
            await Task.Delay(TimeSpan.FromSeconds(30), stoppingToken);

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    using (var scope = _serviceProvider.CreateScope())
                    {
                        var syncService = scope.ServiceProvider.GetRequiredService<ISyncService>();
                        
                        // Obtener configuración
                        var config = syncService.ObtenerConfiguracion();

                        // Procesar archivos de productos (Sucursales)
                        if (config.ProcesaProductos && config.GeneraScriptSync)
                        {
                            var rutaIntercambio = config.RutaIntercambio;

                            if (!string.IsNullOrEmpty(rutaIntercambio) && Directory.Exists(rutaIntercambio))
                            {
                                _logger.LogInformation("Iniciando procesamiento automático de sincronización de productos...");
                                
                                var resultado = await syncService.ProcesarArchivosSyncAsync(rutaIntercambio);

                                if (resultado.Exitoso)
                                {
                                    if (resultado.ArchivosProcessados > 0)
                                    {
                                        _logger.LogInformation(
                                            "Sincronización de productos completada: {Archivos} archivos, {Registros} registros sincronizados",
                                            resultado.ArchivosProcessados,
                                            resultado.RegistrosSincronizados);
                                    }
                                }
                                else
                                {
                                    _logger.LogWarning(
                                        "Sincronización de productos completada con errores: {Mensaje}. Errores: {Errores}",
                                        resultado.Mensaje,
                                        resultado.Errores);
                                    
                                    foreach (var error in resultado.DetalleErrores)
                                    {
                                        _logger.LogError("Error de sincronización: {Error}", error);
                                    }
                                }
                            }
                            else
                            {
                                _logger.LogWarning("Ruta de intercambio de productos no válida o no existe: {Ruta}", rutaIntercambio);
                            }
                        }
                        else if (config.GeneraProductos)
                        {
                            // Si genera archivos de productos (Central), no hacer nada (solo log en la primera ejecución)
                            _logger.LogDebug("Esta instalación genera archivos de productos (Central), no se procesan archivos de sincronización de productos");
                        }

                        // Procesar archivos de remitos (Central)
                        if (config.ProcesaRemitos && config.GeneraScriptSync)
                        {
                            var rutaIntercambioRemitos = config.RutaIntercambioRemitos;

                            if (!string.IsNullOrEmpty(rutaIntercambioRemitos) && Directory.Exists(rutaIntercambioRemitos))
                            {
                                _logger.LogInformation("Iniciando procesamiento automático de sincronización de remitos...");
                                
                                // Obtener número de usuario activo (debe implementarse según el sistema de autenticación)
                                var numUsuar = 0; // Valor por defecto, debe obtenerse del usuario autenticado
                                
                                // PRIMERO: Procesar archivos TXT y convertirlos a JSON
                                var archivosTxt = Directory.GetFiles(rutaIntercambioRemitos, "*.txt");
                                foreach (var archivoTxt in archivosTxt)
                                {
                                    _logger.LogInformation("Procesando archivo TXT de remito: {Archivo}", Path.GetFileName(archivoTxt));
                                    await syncService.GenerarArchivoSyncRemitoDesdeTxtAsync(archivoTxt, numUsuar);
                                }
                                
                                // SEGUNDO: Procesar archivos JSON de remitos
                                var resultado = await syncService.ProcesarArchivosRemitosAsync(rutaIntercambioRemitos, numUsuar);

                                if (resultado.Exitoso)
                                {
                                    if (resultado.ArchivosProcessados > 0)
                                    {
                                        _logger.LogInformation(
                                            "Sincronización de remitos completada: {Archivos} archivos, {Registros} movimientos sincronizados",
                                            resultado.ArchivosProcessados,
                                            resultado.RegistrosSincronizados);
                                    }
                                }
                                else
                                {
                                    _logger.LogWarning(
                                        "Sincronización de remitos completada con errores: {Mensaje}. Errores: {Errores}",
                                        resultado.Mensaje,
                                        resultado.Errores);
                                    
                                    foreach (var error in resultado.DetalleErrores)
                                    {
                                        _logger.LogError("Error de sincronización de remitos: {Error}", error);
                                    }
                                }
                            }
                            else
                            {
                                _logger.LogWarning("Ruta de intercambio de remitos no válida o no existe: {Ruta}", rutaIntercambioRemitos);
                            }
                        }
                        else if (config.GeneraRemitos)
                        {
                            // Si genera archivos de remitos (Sucursal Neuquén), no hacer nada (solo log en la primera ejecución)
                            _logger.LogDebug("Esta instalación genera archivos de remitos (Sucursal Neuquén), no se procesan archivos de sincronización de remitos");
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error en el procesamiento automático de sincronización");
                }

                // Esperar el intervalo configurado antes de la siguiente ejecución
                await Task.Delay(_intervalo, stoppingToken);
            }

            _logger.LogInformation("SyncBackgroundService detenido");
        }
    }
}
