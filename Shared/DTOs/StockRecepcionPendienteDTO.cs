using System;

namespace IMPLANPROD.Shared.DTOs
{
    public sealed class StockRecepcionPendienteDTO
    {
        public int Numebore { get; set; }
        public int NumProv { get; set; }
        public string Proveedor { get; set; } = string.Empty;
        public decimal CantPendiente { get; set; }
        public string RemitoProveedor { get; set; } = string.Empty;
        public DateTime FechaRecepcion { get; set; }
    }
}
