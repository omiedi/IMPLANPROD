namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para sincronización de la tabla Formulas entre Central y Sucursal
    /// </summary>
    public class SyncFormulasDTO
    {
        public string Tabla { get; set; } = "FORMULAS";
        public string Operacion { get; set; } = "UPSERT";
        public DateTime FechaGeneracion { get; set; }
        public int CodIConj { get; set; }
        public List<FormulaItemDTO> Componentes { get; set; } = new();
    }

    /// <summary>
    /// Representa un componente individual de la fórmula
    /// </summary>
    public class FormulaItemDTO
    {
        public int? CodIConj { get; set; }
        public short? CodiEmpr { get; set; }
        public string? CodXConj { get; set; }
        public int? CodiElem { get; set; }
        public string? CodXElem { get; set; }
        public short? NuorIt { get; set; }
        public decimal? UsoBruto { get; set; }
        public decimal? Unis_Form { get; set; }
        public decimal? Lar_Corte { get; set; }
        public int? Pladispo { get; set; }
        public string? RefeComp { get; set; }
        public short? Marborra { get; set; }
        public string? UMED_FORM { get; set; }
        public string? COMPONENTE { get; set; }
        public short? MODUXBB { get; set; }
        public decimal? CORTE { get; set; }
        public string? VALOBADA_1 { get; set; }
    }
}
