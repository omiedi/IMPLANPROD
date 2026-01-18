using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.Entities
{
    public class CierreParcial
    {
        public int IdCierre { get; set; }              // Clave primaria
        public int IdSubor { get; set; }               // Orden de fabricación
        public DateTime FechaCierre { get; set; }      // Fecha del cierre
        public int CantidadFabricada { get; set; }     // Cantidad informada
        public string Usuario { get; set; }            // Usuario que realizó el cierre
        public string Estado { get; set; }             // 'activo' o 'revertido'
    }

}
//CREATE TABLE cierres_parciales (
//    id_cierre INT PRIMARY KEY IDENTITY(1,1),
//    IdSubor INT NOT NULL,                     -- ID de la orden de fabricación
//    fecha_cierre DATETIME NOT NULL,
//    cantidad_fabricada INT NOT NULL,
//    usuario NVARCHAR(100) NOT NULL,
//    estado NVARCHAR(20) NOT NULL              -- 'activo' o 'revertido'
//);