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
        /// Procesa todos los archivos JSON de remitos pendientes en la carpeta de intercambio (Central)
        /// Genera INSERT en la tabla MOVISTO
        /// </summary>
        Task<SyncResultDTO> ProcesarArchivosRemitosAsync(string rutaIntercambio, int numUsuar);

        /// <summary>
        /// Genera un archivo JSON de sincronización de remitos desde un archivo TXT
        /// Se usa para convertir archivos TXT del sistema viejo a JSON
        /// </summary>
        Task<bool> GenerarArchivoSyncRemitoDesdeTxtAsync(string rutaArchivoTxt, int numUsuar);

        /// <summary>
        /// Procesa todos los archivos JSON de órdenes de reparación pendientes en la carpeta de intercambio (Central)
        /// Genera INSERT en la tabla MOVISTO
        /// </summary>
        Task<SyncResultDTO> ProcesarArchivosReparacionesAsync(string rutaIntercambio, int numUsuar);

        /// <summary>
        /// Genera un archivo JSON de sincronización de órdenes de reparación desde un archivo TXT
        /// Se usa para convertir archivos TXT del sistema viejo a JSON
        /// </summary>
        Task<bool> GenerarArchivoSyncReparacionDesdeTxtAsync(string rutaArchivoTxt, int numUsuar);

        /// <summary>
        /// Procesa todos los archivos JSON de ajustes de stock pendientes en la carpeta de intercambio (Central)
        /// Genera INSERT en la tabla MOVISTO
        /// </summary>
        Task<SyncResultDTO> ProcesarArchivosAjustesAsync(string rutaIntercambio, int numUsuar);

        /// <summary>
        /// Genera un archivo JSON de sincronización de ajustes de stock desde un archivo TXT
        /// Se usa para convertir archivos TXT del sistema viejo a JSON
        /// </summary>
        Task<bool> GenerarArchivoSyncAjusteDesdeTxtAsync(string rutaArchivoTxt, int numUsuar);

        /// <summary>
        /// Genera un archivo JSON de sincronización de remitos directamente desde una lista de movimientos
        /// Se usa cuando IMPLANPROD genera el remito con múltiples items (sin archivo TXT)
        /// </summary>
        Task<bool> GenerarArchivoSyncRemitoDesdeDatosAsync(List<RemitoItemDTO> movimientos, int numUsuar);

        /// <summary>
        /// Genera un archivo JSON de sincronización de remitos directamente desde datos de movimiento
        /// Se usa cuando IMPLANPROD genera el remito (sin archivo TXT)
        /// Sobrecarga para compatibilidad con casos de un solo item
        /// </summary>
        Task<bool> GenerarArchivoSyncRemitoDesdeDatosAsync(RemitoItemDTO movimiento, int numUsuar);

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

        /// <summary>
        /// Lee un archivo Ocompra_Sucursal.txt y lo convierte a DTO
        /// Este archivo se genera en la sucursal para sincronizar pedidos de reposición con la central
        /// </summary>
        Task<OcompraSucursalArchivoResultadoDTO> LeerArchivoOcompraSucursalAsync(string rutaArchivo);

        /// <summary>
        /// Procesa los registros de un archivo Ocompra_Sucursal.txt
        /// Inserta o actualiza los registros en la tabla OC_SUCURSAL
        /// </summary>
        Task<bool> ProcesarRegistrosOcompraSucursalAsync(OcompraSucursalArchivoResultadoDTO archivoResultado, int numUsuar);

        /// <summary>
        /// Procesa todos los archivos Ocompra_Sucursal*.txt en la ruta de intercambio
        /// Este método se llama desde la central para procesar archivos enviados por la sucursal
        /// </summary>
        Task<SyncResultDTO> ProcesarArchivosOcompraSucursalAsync(string rutaIntercambio, int numUsuar);

        /// <summary>
        /// Genera el archivo ARCHIVO_GENERO_PEDIDO después de generar un pedido
        /// Este archivo se envía a la sucursal para notificar que el pedido fue generado
        /// Basado en código VB6: GENERAR_ARCHIVO_GENERO_PEDIDO
        /// </summary>
        Task<bool> GenerarArchivoGeneroPedidoAsync(int nroPedido, int nroClie, int nroCompra, string accion, int numUsuar);
    }
}
