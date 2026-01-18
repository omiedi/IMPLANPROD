using System;

namespace IMPLANPROD.Shared.DTOs
{
    public class RecepcionLibreDTO
    {
        public string NumeroRemito { get; set; }
        public int NumeroProveedor { get; set; }
        public string RazonSocialProveedor { get; set; }
        public int CodigoInterno { get; set; }
        public string CodigoExterno { get; set; }
        public string Descripcion { get; set; }
        public string Unidad { get; set; }
        public decimal CantidadNominal { get; set; }
        public int Deposito { get; set; }
        public string NombreDeposito { get; set; } = "";
    }
}
