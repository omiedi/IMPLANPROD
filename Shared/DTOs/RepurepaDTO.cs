namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mostrar los repuestos de una reparación (REPUPERA) con datos del MASTER.
    /// 
    /// Nota:
    /// - En VB6 se resolvía CODIGO$ = COD_EXTE(CI%) y DESCRIPCION$ = DESCRIP$(CI%)
    /// - Acá lo devolvemos ya resuelto desde el backend para simplificar el frontend.
    /// </summary>
    public class RepurepaDTO
    {
        public int Id { get; set; }

        public int NumRepa { get; set; }

        public int Cod_Inte { get; set; }

        public decimal Cant_Sol { get; set; }

        public string? CodigoExterno { get; set; }

        public string? Descripcion { get; set; }

        public string? UnidadMedida { get; set; }

        public int? Origen { get; set; }

        public short? MarBorra { get; set; }

        public short? Nuord_Item { get; set; }
    }
}
