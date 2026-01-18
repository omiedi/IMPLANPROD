using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad REPUESRE - Repuestos por Reparación
    /// Almacena la relación entre reparaciones y los repuestos necesarios con su uso unitario
    /// </summary>
    [Table("REPUESRE")]
    public class Repuesre : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Código de reparación (vincula con INDIREP)
        /// </summary>
        [Display(Name = "Código Reparación")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(16, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string CodigoReparacion { get; set; } = null!;

        /// <summary>
        /// Código interno del producto (vincula con MASTER.Codint)
        /// </summary>
        [Display(Name = "Código Interno")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public short CodigoInterno { get; set; }

        /// <summary>
        /// Uso unitario del repuesto para esta reparación
        /// </summary>
        [Display(Name = "Uso Unitario")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public decimal Uso { get; set; }

        /// <summary>
        /// Estado del registro
        /// </summary>
        public short? Status { get; set; }
        
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
/*
 * CREATE TABLE [dbo].[REPUESRE](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [CodigoReparacion] [nvarchar](16) NOT NULL,
    [CodigoInterno] [smallint] NOT NULL,
    [Uso] [decimal] NOT NULL,
    [Status] [smallint] NULL,
    [AddRecord] [datetime2](7) NULL,
    [LastUpdate] [datetime2](7) NULL,
    [CodiEmprNet] [int] NULL,
    [NumUsuar] [int] NULL,
 CONSTRAINT [PK_REPUESRE] PRIMARY KEY CLUSTERED 
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE UNIQUE NONCLUSTERED INDEX [IX_REPUESRE_CodigoReparacion_CodigoInterno] ON [dbo].[REPUESRE]
(
    [CodigoReparacion] ASC,
    [CodigoInterno] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO

 */