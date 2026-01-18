using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.DTOs
{
    public class MasterListadoIndexDTO
    {
        public string? Id { get; set; }
        public string? Codint { get; set; }          // Código Interno
        public string? Codigo { get; set; }          // Código
        public string? Descripcion { get; set; }     // Descripción del producto
        public string? Umedida { get; set; }         // Unidad de medida
        public string? Timate { get; set; }              // Código de tipo de material (clave)
        public string? TipoMaterialDescripcion { get; set; } // Descripción obtenida desde Tatimat
    }
}
