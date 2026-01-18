using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.DTOs
{
   
        public class Master1DTO
        {
            public int Codint { get; set; } = 0;        // Para el Id/codint
            public string Codigo { get; set; } = string.Empty; // Para el código externo
            public string Descripcion { get; set; } = string.Empty;
        }
    
}
