using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa la tabla FLEXCRL para control de procesos de negocio
    /// </summary>
    [Table("FLEXCRL")]
    public class FLEXCRL : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador único de la tabla (clave primaria autoincremental)
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        /// <summary>
        /// Proceso de negocio asociado al control
        /// </summary>
        [Column("PROBUS")]
        [StringLength(32)]
        public string? PROBUS { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public short? CodiEmpr { get; set; }

        /// <summary>
        /// Clasificación del proceso de negocio
        /// </summary>
        [Column("CLABUS")]
        [StringLength(32)]
        public string? CLABUS { get; set; }

        /// <summary>
        /// Valor de control del proceso (almacenado como texto extenso)
        /// </summary>
        [Column("VALCONTROL", TypeName = "ntext")]
        public string? VALCONTROL { get; set; }

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

// Estructura de tabla SQL original:
// CREATE TABLE [dbo].[FLEXCRL](
//     [PROBUS] [nvarchar](32) NULL,
//     [CodiEmpr] [smallint] NULL,
//     [CLABUS] [nvarchar](32) NULL,
//     [VALCONTROL] [ntext] NULL,
//     [ID] [int] IDENTITY(1,1) NOT NULL,
//  CONSTRAINT [PK_FLEXCRL] PRIMARY KEY CLUSTERED 
// (
//     [ID] ASC
// )
