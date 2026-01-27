using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Interfaz para el servicio de sincronización entre Central y Sucursal
    /// </summary>
    public interface ISyncService
    {
        /// <summary>
        /// Genera archivo JSON de sincronización para un registro Master (Central)
        /// </summary>
        Task<bool> GenerarArchivoSyncMasterAsync(int codint);

        /// <summary>
        /// Genera archivo JSON de sincronización para un conjunto de Formulas (Central)
        /// </summary>
        Task<bool> GenerarArchivoSyncFormulasAsync(int codiConj);

        /// <summary>
        /// Genera archivo JSON de sincronización para un conjunto de SecOper (Central)
        /// </summary>
        Task<bool> GenerarArchivoSyncSecOperAsync(int codiConj);

        /// <summary>
        /// Procesa todos los archivos de sincronización pendientes en la carpeta de intercambio (Sucursal)
        /// </summary>
        Task<SyncResultDTO> ProcesarArchivosSyncAsync(string rutaIntercambio);

        /// <summary>
        /// Registra el procesamiento de un archivo de sincronización
        /// </summary>
        Task RegistrarProcesamientoAsync(string archivo, string tabla, string operacion, string estado, string? error = null, int? codintAfectado = null);

        /// <summary>
        /// Obtiene la configuración de sincronización desde appsettings
        /// </summary>
        SyncConfigDTO ObtenerConfiguracion();

        /// <summary>
        /// Guarda la configuración de sincronización en appsettings
        /// </summary>
        Task<bool> GuardarConfiguracionAsync(SyncConfigDTO config);
    }
}
