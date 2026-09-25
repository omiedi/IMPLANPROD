namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Request para registrar la anotación automática "Items Remitidos" en el
    /// anotador de pedidos cuando se genera un remito de cliente.
    /// Solo viaja el número de remito largo (MOVISTO.DoPriLar, ej. "RT-X-0001-00006449"):
    /// el server determina los ítems remitidos por pedido leyendo MOVISTO
    /// (CodiMovi = "RT" y NuPedCli &gt; 0), por lo que los ítems del remito que no
    /// están asociados a un pedido quedan automáticamente excluidos.
    /// </summary>
    public class RegistrarItemsRemitidosDTO
    {
        /// <summary>Documento primario largo del remito (MOVISTO.DoPriLar)</summary>
        public string RemitoLargo { get; set; } = string.Empty;
    }
}
