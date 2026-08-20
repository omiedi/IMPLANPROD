namespace IMPLANPROD.Shared.DTOs
{
    public class BrecepenPreviewDTO
    {
        public string Fecha { get; set; } = "";
        public int CodInte { get; set; }
        public string IdLote { get; set; } = "";
        public string CodMovi { get; set; } = "";
        public string DopriCor { get; set; } = "";
        public string DoseCor { get; set; } = "";
        public string ReferCor { get; set; } = "";
        public short DepoMovi { get; set; }
        public short NumeProv { get; set; }
        public double CantIngre { get; set; }
        public double CantSalid { get; set; }
        public decimal Saldo { get; set; }
        public bool Migrar { get; set; }
    }
}
