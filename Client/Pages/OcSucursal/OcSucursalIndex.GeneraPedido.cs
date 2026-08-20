using IMPLANPROD.Client.Repositories;
using IMPLANPROD.Shared.DTOs;
using CurrieTechnologies.Razor.SweetAlert2;

namespace IMPLANPROD.Client.Pages.OcSucursal
{
    /// <summary>
    /// Componente que maneja la lógica de generación de pedidos desde órdenes de compra de sucursal
    /// </summary>
    public partial class OcSucursalIndex
    {
        /// <summary>
        /// Genera un pedido a partir de la orden de compra de sucursal seleccionada
        /// </summary>
        private async Task GenerarPedido()
        {
            if (itemSeleccionado == null)
            {
                return;
            }

            try
            {
                // Confirmar generación de pedido
                var result = await sweetAlertService.FireAsync(new SweetAlertOptions
                {
                    Title = "¿Generar Pedido?",
                    Text = $"¿Desea generar el pedido para la orden de compra {itemSeleccionado.OCompraSuc} de la sucursal {itemSeleccionado.Sucursal}?",
                    Icon = SweetAlertIcon.Question,
                    ShowCancelButton = true,
                    ConfirmButtonText = "Sí, Generar",
                    CancelButtonText = "Cancelar"
                });

                if (!result.IsConfirmed)
                {
                    return; // Usuario canceló
                }

                // Crear DTO para generar pedido
                var request = new GenerarPedidoDTO
                {
                    Sucursal = itemSeleccionado.Sucursal.Value,
                    OCompraSuc = itemSeleccionado.OCompraSuc.Value,
                    NcPedido = itemSeleccionado.NclienPedido
                };

                // Llamar al endpoint para generar pedido
                var response = await repository.Post<GenerarPedidoDTO, int>("api/OcSucursal/generar-pedido", request);

                if (!response.Error && response.Response != null)
                {
                    await sweetAlertService.FireAsync("Éxito", $"Pedido generado correctamente. Número de pedido: {response.Response}", SweetAlertIcon.Success);
                    
                    // Cerrar modal y recargar datos
                    CerrarModalDetalle();
                    await CargarDatos();
                }
                else
                {
                    await sweetAlertService.FireAsync("Error", await response.GetErrorMessageAsync(), SweetAlertIcon.Error);
                }
            }
            catch (Exception ex)
            {
                await sweetAlertService.FireAsync("Error", $"Error al generar pedido: {ex.Message}", SweetAlertIcon.Error);
            }
        }
    }
}
