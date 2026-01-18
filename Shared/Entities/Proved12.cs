using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa la tabla Proved12.
    /// Todos los comentarios y nombres siguen el estándar del proyecto y están en español.
    /// </summary>
    public class Proved12 : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador único autoincremental para Proved12 (nuevo campo IDENTITY).
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Número de cliente (se mantiene como int, pero ya no es identity).
        /// </summary>
        [Column("Nume_Cli")]
        public int Nume_Cli { get; set; }

        public decimal? Lcre_Cli { get; set; }
        public int? CodiEmpr { get; set; }
        public decimal? Sald_Cli { get; set; }
        public short? CuMa_Cli { get; set; }
        public short? GrCo_Cli { get; set; }

        [MaxLength(int.MaxValue)]
        public string? Nota_Cli { get; set; }

        [MaxLength(10)]
        public string? C_Ret { get; set; }

        [MaxLength(10)]
        public string? Posivat_Cli { get; set; }

        [MaxLength(12)]
        public string? Codi_Cli { get; set; }

        [Required]
        [MaxLength(128)]
        public string? Raso_Cli { get; set; }

        [MaxLength(32)]
        public string? Fant_Cli { get; set; }

        public short? Stat_Cli { get; set; }

        [MaxLength(64)]
        public string? Domi_Cli { get; set; }

        [MaxLength(16)]
        public string? Cpos_Cli { get; set; }

        [MaxLength(48)]
        public string? Loca_Cli { get; set; }

        [MaxLength(48)]
        public string? Pais_Cli { get; set; }

        [MaxLength(48)]
        public string? Prov_Cli { get; set; }

        [MaxLength(32)]
        public string? Tele_Cli { get; set; }

        [MaxLength(32)]
        public string? Tele2_Cli { get; set; }

        [MaxLength(32)]
        public string? Faxi_Cli { get; set; }

        [MaxLength(64)]
        public string? Mail_Cli { get; set; }

        public short? Piva_Cli { get; set; }

        [MaxLength(24)]
        public string? Cuit_Cli { get; set; }

        [MaxLength(24)]
        public string? Nibr_Cli { get; set; }

        [MaxLength(24)]
        public string? Prvd_Cli { get; set; }

        [MaxLength(48)]
        public string? Ctac_Cli { get; set; }

        [MaxLength(32)]
        public string? Hora_Cli { get; set; }

        [MaxLength(48)]
        public string? Ord_Cheq { get; set; }

        public short? Cpag_Cli { get; set; }

        [MaxLength(int.MaxValue)]
        public string? CCom_Cli { get; set; }

        public short? Lpre_Cli { get; set; }
        public short? Vend_Cli { get; set; }
        public short? Tcpr_Cli { get; set; }

        [MaxLength(2)]
        public string? CaIb_Cli { get; set; }

        [MaxLength(32)]
        public string? Dsta_Cli { get; set; }

        public decimal? Calinteg { get; set; }

        [MaxLength(int.MaxValue)]
        public string? ObsCompras { get; set; }

        [MaxLength(int.MaxValue)]
        public string? LuEntCompras { get; set; }

        public short? CategoriaIBCABA_Cli { get; set; }
        public short? CategoriaIBMISIONES_Cli { get; set; }
        public short? CategoriaIBTUCU_Cli { get; set; }

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
