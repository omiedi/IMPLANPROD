using IMPLANPROD.Server.Services;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio en segundo plano que procesa automáticamente archivos de sincronización
    /// Solo se ejecuta en instalaciones que NO son Central (Sucursales)
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

                        // Solo procesar si NO es Central y la sincronización está habilitada
                        if (!config.EsCentral && config.GeneraScriptSync)
                        {
                            var rutaIntercambio = config.RutaIntercambio;

                            if (!string.IsNullOrEmpty(rutaIntercambio) && Directory.Exists(rutaIntercambio))
                            {
                                _logger.LogInformation("Iniciando procesamiento automático de sincronización...");
                                
                                var resultado = await syncService.ProcesarArchivosSyncAsync(rutaIntercambio);

                                if (resultado.Exitoso)
                                {
                                    if (resultado.ArchivosProcessados > 0)
                                    {
                                        _logger.LogInformation(
                                            "Sincronización completada: {Archivos} archivos, {Registros} registros sincronizados",
                                            resultado.ArchivosProcessados,
                                            resultado.RegistrosSincronizados);
                                    }
                                }
                                else
                                {
                                    _logger.LogWarning(
                                        "Sincronización completada con errores: {Mensaje}. Errores: {Errores}",
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
                                _logger.LogWarning("Ruta de intercambio no válida o no existe: {Ruta}", rutaIntercambio);
                            }
                        }
                        else if (config.EsCentral)
                        {
                            // Si es Central, no hacer nada (solo log en la primera ejecución)
                            _logger.LogDebug("Esta instalación es Central, no se procesan archivos de sincronización");
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
