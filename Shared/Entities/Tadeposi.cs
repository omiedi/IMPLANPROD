using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    public class Tadeposi : ITimestampEntity, IAuditableEntity
    {
        public int Id { get; set; }

        [Display(Name = "Código Depósito")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public int CodigoDepo { get; set; }

        [Display(Name = "Descripción")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(64, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? Descripcion { get; set; }

        [Display(Name = "Descripción Abreviada")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(12, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? DescripAbreviada { get; set; }

        public int? CodPlanta { get; set; }

        public bool disponible { get; set; } = true;
         
        [Display(Name = "Estado")]
        public int Status { get; set; } = 0;

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}
