using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("INDIREP")]
    public class Indirep : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Display(Name = "Código")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(24, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? Codigo { get; set; }

        [Display(Name = "Descripción")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(128, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Estado del pedido
        /// </summary>
        public short? Status { get; set; }
        
        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// Se establece automáticamente cuando se crea el movimiento de stock
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se actualiza automáticamente cada vez que se modifica el movimiento
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// Se actualiza automáticamente con la empresa del usuario actual
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// Se actualiza automáticamente con el ID del usuario actual
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}
/*
 * CREATE TABLE [dbo].[INDIREP](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [Codigo] [nvarchar](24) NOT NULL,
    [Descripcion] [nvarchar](128) NOT NULL,
    [Status] [smallint] NULL,
    [AddRecord] [datetime2](7) NULL,
    [LastUpdate] [datetime2](7) NULL,
    [CodiEmprNet] [int] NULL,
    [NumUsuar] [int] NULL,
 CONSTRAINT [PK_INDIREP] PRIMARY KEY CLUSTERED 
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE UNIQUE NONCLUSTERED INDEX [IX_INDIREP_Codigo] ON [dbo].[INDIREP]
(
    [Codigo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO

 */