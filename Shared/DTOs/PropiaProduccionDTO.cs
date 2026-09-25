using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para la pantalla de Producción Propia.
    /// Representa los movimientos de MOVISTO donde CantIngre > 0, CodiMovi = 'pf' y DoPriLar LIKE 'OF-%'.
    /// Cada registro corresponde a una fabricación (ingreso de producto terminado).
    /// </summary>
    public class PropiaProduccionDTO
    {
        /// <summary>
        /// Fecha del movimiento (MOVISTO.FechMov)
        /// </summary>
        public DateTime? FechMov { get; set; }

        /// <summary>
        /// Documento largo (MOVISTO.DoPriLar) — ej: "OF-5604"
        /// </summary>
        public string? DoPriLar { get; set; }

        /// <summary>
        /// Código externo del producto (MOVISTO.Cod_Exte)
        /// </summary>
        public string? Cod_Exte { get; set; }

        /// <summary>
        /// Descripción del producto (MOVISTO.Descrip)
        /// </summary>
        public string? Descrip { get; set; }

        /// <summary>
        /// Cantidad ingresada / fabricada (MOVISTO.CantIngre)
        /// </summary>
        public decimal? CantIngre { get; set; }
    }

    /// <summary>
    /// DTO para los consumos de una producción propia.
    /// Representa los movimientos de MOVISTO donde CantSalid > 0, CodiMovi = 'cf' y DoSecLar = 'pf-XXXX'.
    /// Cada registro corresponde a un componente consumido para fabricar el producto terminado.
    /// </summary>
    public class ConsumoProduccionDTO
    {
        /// <summary>
        /// Fecha del movimiento (MOVISTO.FechMov)
        /// </summary>
        public DateTime? FechMov { get; set; }

        /// <summary>
        /// Código externo del componente consumido (MOVISTO.Cod_Exte)
        /// </summary>
        public string? Cod_Exte { get; set; }

        /// <summary>
        /// Descripción del componente consumido (MOVISTO.Descrip)
        /// </summary>
        public string? Descrip { get; set; }

        /// <summary>
        /// Cantidad consumida / salida (MOVISTO.CantSalid)
        /// </summary>
        public decimal? CantSalid { get; set; }
    }
}
