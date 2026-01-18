using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.DTOs
{
    public class MovistoDTO
    {
        public DateTime FeReMovi { get; set; }
        public DateTime FechMov { get; set; }

        public string? CodiMovi { get; set; }
        public int? Cod_Inte { get; set; }
        public string? Cod_exte { get; set; }
        public int? DepoMovi { get; set; }
        public decimal? CantIngre { get; set; }
        public decimal? CantSalid { get; set; }
        public string? ReferCor { get; set; }
        public string? IdenLote { get; set; }
        public string? DoPriCor { get; set; }
        public string? DoPriLar { get; set; }
        public string? DoSecCor { get; set; }
        public string? DoSecLar { get; set; }
        public int? Nume_Prov { get; set; }
        public string? Descrip { get; set; }
    }
}
