namespace IMPLANPROD.Shared.DTOs
{
    public sealed class PedidoVentaPendienteDetalleDTO
    {
        public int NroPed { get; set; }
        // N° de cliente y razón social provenientes de PEDDETA (NroClie / RaSoClie).
        // Se incluyen en el reporte "Pedidos Pendientes" para identificar al cliente
        // asociado a cada ítem pendiente, igual que en el sistema VB6 original.
        public int NroClie { get; set; }
        public string Rasoclie { get; set; } = string.Empty;
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public decimal CantSolicitada { get; set; }
        public string UMedida { get; set; } = string.Empty;
        public decimal CantEntregada { get; set; }
        public decimal Saldo { get; set; }
        public DateTime? FechaEntrega { get; set; }
    }
}
