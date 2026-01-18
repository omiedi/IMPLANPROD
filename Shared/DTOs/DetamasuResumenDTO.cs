using System;

namespace IMPLANPROD.Shared.DTOs
{
    public class DetamasuResumenDTO
    {
        public int Cod_Inte { get; set; }
        public string? Cod_Exte { get; set; }
        public string? Codigo { get; set; }
        public string? Descripcion { get; set; }

        // BR: ingresos (CantIngre)
        public decimal Br { get; set; }

        // CF: consumos (CantSalid)
        public decimal Cf { get; set; }

        // DV: devoluciones al cliente (CantSalid)
        public decimal Dv { get; set; }

        // AJ: ajustes manuales netos (CantIngre - CantSalid) cuando CodiMovi = 'AJ'
        public decimal Aj { get; set; }

        // RESUL = BR - (CF + DV) + AJ
        public decimal Resul { get; set; }

        // Compatibilidad: algunas pantallas usaban una única columna "Cantidad".
        // En Materiales, esta Cantidad representa el resultado (RESUL).
        public decimal Cantidad { get; set; }
        public DateTime? FechaUltMov { get; set; }
    }
}
