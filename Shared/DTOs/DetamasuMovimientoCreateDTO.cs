using System;

namespace IMPLANPROD.Shared.DTOs
{
    public class DetamasuMovimientoCreateDTO
    {
        public string? CodiMovi { get; set; }
        public int? Cod_Inte { get; set; }
        public int? Talon { get; set; }
        public int? Nume_Clie { get; set; }
        public decimal Cantidad { get; set; }
        public DateTime? FechMov { get; set; }
        public string? Referencia { get; set; }
    }
}
