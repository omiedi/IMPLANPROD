using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa los vendedores en el sistema.
    /// Mapea a la tabla VENDEDOR en la base de datos.
    /// </summary>
    [Table("VENDEDOR")]
    public class Vendedor : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador único del vendedor (campo IDENTITY)
        /// </summary>
        [Key]
        [Column("Id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Número de vendedor
        /// </summary>
        public int? NumVendedor { get; set; }

        /// <summary>
        /// Nombre del vendedor
        /// </summary>
        [StringLength(48)]
        public string? NombreVendedor { get; set; }

        /// <summary>
        /// Comisión por venta (porcentaje)
        /// </summary>
        [Column(TypeName = "decimal(18,2)")]
        public decimal? ComVenta { get; set; }

        /// <summary>
        /// Comisión por cobranza (porcentaje)
        /// </summary>
        [Column(TypeName = "decimal(18,2)")]
        public decimal? ComCobranza { get; set; }

        /// <summary>
        /// Estado del vendedor (0=activo, -1=inactivo)
        /// </summary>
        public int? Status { get; set; }

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
-- Instrucciones SQL para crear la tabla VENDEDOR en SQL Server Management Studio:

-- Crear la tabla VENDEDOR
CREATE TABLE [dbo].[VENDEDOR](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NumVendedor] [int] NULL,
    [NombreVendedor] [nvarchar](48) NULL,
    [ComVenta] [decimal](18,2) NULL,
    [ComCobranza] [decimal](18,2) NULL,
    [Status] [int] NULL,
    CONSTRAINT [PK_VENDEDOR] PRIMARY KEY CLUSTERED ([Id] ASC)
) ON [PRIMARY]

-- Crear índice único en NumVendedor (opcional, si se requiere que sea único)
-- CREATE UNIQUE INDEX [IX_VENDEDOR_NumVendedor] ON [dbo].[VENDEDOR] ([NumVendedor] ASC) WHERE [NumVendedor] IS NOT NULL

-- Insertar algunos datos de ejemplo (opcional)
-- INSERT INTO [dbo].[VENDEDOR] ([NumVendedor], [NombreVendedor], [ComVenta], [ComCobranza], [Status]) VALUES 
-- (1, 'Juan Pérez', 5.00, 2.50, 0),
-- (2, 'María García', 4.50, 2.00, 0),
-- (3, 'Carlos López', 6.00, 3.00, 0),
-- (4, 'Ana Martínez', 5.50, 2.75, 0)
*/
