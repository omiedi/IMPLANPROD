using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.DTOs
{
    public class RegistroOFabricacionDTO
    {
        public int CodInte { get; set; }
        public string? DoPriCor { get; set; }
        public string? Codigo { get; set; }
        public string? Descripcion { get; set; }
        public decimal? Ajuste { get; set; }
        public string? IdText { get; set; }
        public DateTime? FechaMovimiento { get; set; }
        public int? Deposito { get; set; }
        public string? NombreDeposito { get; set; }
    }


}
