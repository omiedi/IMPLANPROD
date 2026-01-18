using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa la tabla Formulas en la base de datos
    /// </summary>
    [Table("Formulas")]
    public class Formula
    {
        /// <summary>
        /// Identificador único de la fórmula
        /// </summary>
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Código interno del conjunto (producto a fabricar)
        /// </summary>
        [Column("CodIConj")]
        public int CodIConj { get; set; }

        /// <summary>
        /// Código interno del elemento (componente)
        /// </summary>
        [Column("CodIElem")]
        public int CodIElem { get; set; }

        /// <summary>
        /// Código externo del elemento (componente)
        /// </summary>
        [Column("CodXElem")]
        public string CodXElem { get; set; }

        /// <summary>
        /// Uso bruto del componente en la fórmula
        /// </summary>
        [Column("UsoBruto")]
        public decimal UsoBruto { get; set; }

        /// <summary>
        /// Número de orden del ítem en la fórmula
        /// </summary>
        [Column("NuOrIt")]
        public int NuOrIt { get; set; }
    }
}
