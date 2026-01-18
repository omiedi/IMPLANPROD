using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.DTOs
{
    public class DetalleStockDTO
    {
        public int CI1 { get; set; }
        public string CodXEelem { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public string UM { get; set; } = string.Empty;
        public decimal Uso { get; set; }
        public decimal CantSimulada { get; set; }
        public decimal StAcumulado { get; set; }
        public Dictionary<string, decimal> StockPorDeposito { get; set; } = new Dictionary<string, decimal>();
    }
}
