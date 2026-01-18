using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    public class UnidadMedidas : ITimestampEntity, IAuditableEntity
    {
        // Campo que representa el identificador único de la unidad de medida
        public int Id { get; set; }

        // Campo que representa la descripción abreviada de la unidad de medida
        [Required(ErrorMessage = "La descripción abreviada es obligatoria.")]
        public string? DescripcionAbreviada { get; set; }

        // Campo que representa la descripción completa de la unidad de medida
        [Required(ErrorMessage = "La descripción completa es obligatoria.")]
        public string? Descripcion { get; set; }

        public int? status { get; set; }=0;

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
