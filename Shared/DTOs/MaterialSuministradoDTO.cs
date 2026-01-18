namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mostrar materiales suministrados por el cliente agrupados
    /// Replica la consulta SQL de CARGAMATSUM en VB6
    /// </summary>
    public class MaterialSuministradoDTO
    {
        public int Cod_Inte { get; set; }
        public string? Cod_Exte { get; set; }
        public string? Descripcion { get; set; }
        public decimal Saldo { get; set; } // (SUM(CANTINGRE) - SUM(CANTSALID))
    }

    /// <summary>
    /// DTO para registrar movimientos de material suministrado
    /// Replica la rutina MOVI_SUMINISTRADOS de VB6
    /// </summary>
    public class MovimientoMaterialDTO
    {
        public int Cod_Inte { get; set; }
        public string CodiMovi { get; set; } = string.Empty; // BR, DV, CF
        public decimal Cantidad { get; set; }
        public int Nume_Clie { get; set; }
        public int Talon { get; set; } // NumRepa
        public DateTime FechMov { get; set; }
        public string? Referencia { get; set; }
    }
}
