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
        public string? Observa { get; set; }  // bolrecep.observa
    }

    /// <summary>
    /// Detalle por lote de las recepciones pendientes de aprobación.
    /// A diferencia de <see cref="RecepcionPendienteAprobacionDTO"/> (que agrupa por
    /// recepción), acá se devuelve una fila por cada registro de BOLRECEP: una misma
    /// recepción (mismo NumeroRecepcion) puede tener varias filas, una por lote
    /// (BOLRECEP.IdenLote, ej. "BR-000125-1", "BR-000125-2", ...).
    /// Se usa en el modal "Detalle por Lote" de /recepcion/selproveedor (solapa Aprueba Calidad).
    /// </summary>
    public class RecepcionPendienteAprobacionDetalleDTO
    {
        public int NumeroRecepcion { get; set; }  // bolrecep.numebore
        public int NumeroOrdenCompra { get; set; }  // bolrecep.numoc
        public DateTime FechaRecepcion { get; set; }  // bolrecep.Feyhorecep
        public string NumeroRemito { get; set; }  // numRem_Prov
        public int NumeroProveedor { get; set; }  // bolrecep.numprov
        public string RazonSocial { get; set; }  // Proved12.Raso_Cli
        public string IdenLote { get; set; }  // bolrecep.IdenLote (ej. "BR-000125-1")
        public string Codigo { get; set; }  // bolrecep.CodExte
        public string Descripcion { get; set; }  // bolrecep.DescripItem
        public decimal CantidadNominal { get; set; }  // bolrecep.CantNom
    }
}
