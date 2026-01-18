using System;

namespace IMPLANPROD.Shared.DTOs
{
    public class RecepcionPendienteAprobacionDTO
    {
        public int NumeroRecepcion { get; set; }  // bolrecep.numebore
        public int NumeroOrdenCompra { get; set; }  // bolrecep.numoc
        public DateTime FechaRecepcion { get; set; }  // bolrecep.Feyhorecep
        public string NumeroRemito { get; set; }  // numRem_Prov
        public int NumeroProveedor { get; set; }  // bolrecep.numprov
        public string RazonSocial { get; set; }  // Proveedor.razonsocial
    }
}
