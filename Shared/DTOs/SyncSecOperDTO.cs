namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para sincronización de la tabla SecOper entre Central y Sucursal
    /// </summary>
    public class SyncSecOperDTO
    {
        public string Tabla { get; set; } = "SECOPER";
        public string Operacion { get; set; } = "UPSERT";
        public DateTime FechaGeneracion { get; set; }
        public int CodIConj { get; set; }
        public List<SecOperItemDTO> Operaciones { get; set; } = new();
    }

    /// <summary>
    /// Representa una operación individual de la secuencia de producción
    /// </summary>
    public class SecOperItemDTO
    {
        public int CodIConj { get; set; }
        public short? CodiEmpr { get; set; }
        public string? CodXConj { get; set; }
        public short? NumeOper { get; set; }
        public short? AltOper { get; set; }
        public string? DescOper { get; set; }
        public short? StatOper { get; set; }
        public string? CodMaq0 { get; set; }
        public string? CodMaq1 { get; set; }
        public string? CodMaq2 { get; set; }
        public string? CodMaq3 { get; set; }
        public string? CodMaq4 { get; set; }
        public string? CodMaq5 { get; set; }
        public decimal? HorsPrep { get; set; }
        public decimal? HorsFijo { get; set; }
        public decimal? MinuStan { get; set; }
        public decimal? HorsMHer { get; set; }
        public decimal? HorsLimp { get; set; }
        public decimal? PieCiclo { get; set; }
        public decimal? PorDemor { get; set; }
        public decimal? Cadencia { get; set; }
        public decimal? CoefSolp { get; set; }
        public decimal? LoteStan { get; set; }
        public decimal? CantiOps { get; set; }
        public decimal? Cantiops2 { get; set; }
        public short? CategOps { get; set; }
        public short? Categops2 { get; set; }
        public string? ProgCNum { get; set; }
        public string? HoProFab { get; set; }
        public string? ObseOper { get; set; }
        public short? ProTerOp { get; set; }
        public string? ProvOper { get; set; }
        public decimal? PreUnid { get; set; }
        public short? MonePrec { get; set; }
        public DateTime? FechPrec { get; set; }
        public decimal? Nuoper_Prep { get; set; }
        public short? Catoper_Prep { get; set; }
        public DateTime? Fech_Revis { get; set; }
        public string? Usua_Revis { get; set; }
        public Guid rowguid { get; set; }
        public short? SecuenciaFabricacion { get; set; }
        public short? OperacionTerceros { get; set; }
        public short? NuOrdenManual { get; set; }
    }
}
