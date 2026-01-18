using System;

namespace IMPLANPROD.Shared.DTOs
{
    public class OrdenCompraPendienteDTO
    {
        public decimal CantidadOrdenada { get; set; }  // cant_ocom
        public decimal CantidadRecibida { get; set; }  // Cant_Rec
        public decimal CoeficienteUnidades { get; set; }  // Coef_Unids
        public int NumeroOrdenCompra { get; set; }  // Nume_Ocom
        public DateTime? FechaEntregaSolicitada { get; set; }  // Fentre_Sol
        public int CodigoInterno { get; set; }  // COD_INTE
        public string DocumentoOrigen { get; set; } = string.Empty;  // Doc_Orig
        public int NumeroOrdenItem { get; set; }  // NuOrd_Item
        public decimal NumeroOrdenServicio { get; set; }  // N_ORSERV
        public bool Seleccionado { get; set; }  // Para UI
        public decimal CantidadPendiente => CantidadOrdenada - CantidadRecibida;
        public string Cod_Exte { get; set; } = string.Empty;  // Código externo del repuesto
        public string Descritem { get; set; } = string.Empty;  // Descripción del repuesto
        public decimal CantidadARecibir { get; set; }

        public OrdenCompraPendienteDTO()
        {
            CantidadARecibir = CantidadPendiente;
        }
    }
}
