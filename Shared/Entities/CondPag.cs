using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa las condiciones de pago en el sistema.
    /// Mapea a la tabla CONDPAG en la base de datos.
    /// </summary>
    [Table("CONDPAG")]
    public class CondPag : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador único de la condición de pago (campo IDENTITY)
        /// </summary>
        [Key]
        [Column("Id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Número de condición de pago
        /// </summary>
        public int? NumCondPag { get; set; }

        /// <summary>
        /// Descripción de la condición de pago
        /// </summary>
        [StringLength(48)]
        public string? DescCondPag { get; set; }

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
-- Instrucciones SQL para crear la tabla CONDPAG en SQL Server Management Studio:

-- Crear la tabla CONDPAG
CREATE TABLE [dbo].[CONDPAG](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NumCondPag] [int] NULL,
    [DescCondPag] [nvarchar](48) NULL,
    CONSTRAINT [PK_CONDPAG] PRIMARY KEY CLUSTERED ([Id] ASC)
) ON [PRIMARY]

-- Crear índice único en NumCondPag (opcional, si se requiere que sea único)
-- CREATE UNIQUE INDEX [IX_CONDPAG_NumCondPag] ON [dbo].[CONDPAG] ([NumCondPag] ASC) WHERE [NumCondPag] IS NOT NULL

-- Insertar algunos datos de ejemplo (opcional)
-- INSERT INTO [dbo].[CONDPAG] ([NumCondPag], [DescCondPag]) VALUES 
-- (1, 'Contado'),
-- (2, '30 días'),
-- (3, '60 días'),
-- (4, '90 días')
*/
