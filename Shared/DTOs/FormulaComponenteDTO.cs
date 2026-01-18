using System;
using System.Collections.Generic;
using System.Text;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar un componente de la fórmula de un producto
    /// </summary>
    public class FormulaComponenteDTO
    {
        /// <summary>
        /// Código interno del componente
        /// </summary>
        public int CodIElem { get; set; }

        /// <summary>
        /// Código externo del componente
        /// </summary>
        public string CodXElem { get; set; }

        /// <summary>
        /// Descripción del componente
        /// </summary>
        public string Descripcion { get; set; }

        /// <summary>
        /// Uso bruto del componente en la fórmula
        /// </summary>
        public decimal UsoBruto { get; set; }

        /// <summary>
        /// Stock disponible del componente
        /// </summary>
        public decimal Stock { get; set; }

        /// <summary>
        /// Indica si el código se repite en la fórmula
        /// </summary>
        public bool EsRepetido { get; set; }

        /// <summary>
        /// Necesidad total considerando la cantidad a fabricar (ORDEFABR.CANPROY * FORMULAS.USOBRUTO)
        /// </summary>
        public decimal NecesidadTotal { get; set; }

        /// <summary>
        /// Cantidad a utilizar (editable por el usuario, por defecto igual a NecesidadTotal)
        /// </summary>
        public decimal CantidadAUtilizar { get; set; }
    }
}
