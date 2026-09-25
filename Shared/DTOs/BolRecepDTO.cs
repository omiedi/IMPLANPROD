using System;

namespace IMPLANPROD.Shared.DTOs
{
    public class BolRecepDTO
    {
        public string NameSysTe { get; set; }
        public string CodExpr { get; set; }
        public string Descrip { get; set; }
        public string IdenLote { get; set; }
        public DateTime FechRecep { get; set; }
        public int Cod_Inte { get; set; }
        public string Cod_Este { get; set; }
        public string Description { get; set; }
        public string Unidad { get; set; }
        public decimal CantNom { get; set; }
        public decimal M_Largo { get; set; }
        public decimal M_Ancho { get; set; }
        public int DepoExt { get; set; }
    }

    public class RecepcionMaterialDTO
    {
        public int NumeroOrdenCompra { get; set; }
        public int NumeroOrdenItem { get; set; }
        public string CodigoExterno { get; set; }
        public string Descripcion { get; set; }
        public decimal CantidadRecibida { get; set; }
        public int Deposito { get; set; }
        public string IdentificadorLote { get; set; }
        public string NumeroRemito { get; set; }
        public int NumeroProveedor { get; set; }
        public int Cod_Inte { get; set; }
        public int Numebore { get; set; }
        public decimal? CanToApro { get; set; }
        public int? Status { get; set; }
        public string? Observa { get; set; }


    }
}
