using Microsoft.AspNetCore.Components;

namespace IMPLANPROD.Client.Pages.Stock
{
    /// <summary>
    /// Code-behind para StockIndex.razor
    /// Contiene la lógica de sincronización centralizada para separar responsabilidades
    /// </summary>
    public partial class StockIndex
    {
        // Sincronización
        private bool mostrarModalSync = false;

        /// <summary>
        /// Abre el modal de sincronización
        /// </summary>
        private void AbrirSincronizacion()
        {
            mostrarModalSync = true;
        }

        /// <summary>
        /// Cierra el modal de sincronización
        /// </summary>
        private void CerrarModalSync()
        {
            mostrarModalSync = false;
        }
    }
}
