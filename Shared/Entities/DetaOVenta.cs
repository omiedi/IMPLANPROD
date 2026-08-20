using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad DetaOVenta - Tabla de detalle de órdenes de venta
    /// Basado en VB6 CREACAMPO para DetaOVenta
    /// </summary>
    [Table("DetaOVenta")]
    public class DetaOVenta : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador interno autoincremental (IDENTITY) y clave primaria
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Número de pedido (tipo 4 = int)
        /// </summary>
        [Column("Nroped")]
        public int? Nroped { get; set; }

        /// <summary>
        /// Número de orden de compra (tipo 4 = int)
        /// </summary>
        [Column("Nro_Ocom")]
        public int? NroOcom { get; set; }

        /// <summary>
        /// Código interno del producto (tipo 4 = int)
        /// </summary>
        [Column("Codint")]
        public int? Codint { get; set; }

        /// <summary>
        /// Código externo del producto (tipo 10 = texto, longitud 24)
        /// </summary>
        [Column("CodiExt")]
        [MaxLength(24)]
        public string? CodiExt { get; set; }

        /// <summary>
        /// Descripción del producto (tipo 10 = texto, longitud 64)
        /// </summary>
        [Column("Descrip")]
        [MaxLength(64)]
        public string? Descrip { get; set; }

        /// <summary>
        /// Cantidad pedida (tipo 6 = decimal, longitud 4)
        /// </summary>
        [Column("CanPed", TypeName = "decimal(18,4)")]
        public decimal? CanPed { get; set; }

        /// <summary>
        /// Cantidad entregada (tipo 6 = decimal, longitud 4)
        /// </summary>
        [Column("CantEnt", TypeName = "decimal(18,4)")]
        public decimal? CantEnt { get; set; }

        /// <summary>
        /// Fecha de modificación (tipo 8 = date)
        /// </summary>
        [Column("FechaModif")]
        public DateTime? FechaModif { get; set; }

        /// <summary>
        /// Acción realizada (tipo 10 = texto, longitud 24)
        /// </summary>
        [Column("Accion")]
        [MaxLength(24)]
        public string? Accion { get; set; }

        /// <summary>
        /// Número de cliente del pedido (tipo 4 = int)
        /// </summary>
        [Column("NCtePedido")]
        public int? NCtePedido { get; set; }

        /// <summary>
        /// Marca de borrado (tipo 3 = short)
        /// </summary>
        [Column("MARBORRA")]
        public short? Marborra { get; set; }

        /// <summary>
        /// Código de empresa (valor por defecto 0 según reglas del sistema)
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; } = 0;

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

        // ===================================================================
        // CAMPOS DE AUDITORÍA - Implementación de IAuditableEntity
        // ===================================================================

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
INSTRUCCIONES SQL PARA CREAR LA TABLA:
Ejecutar en SQL Server Management Studio para crear la tabla DetaOVenta:

-- CREAR TABLA DetaOVenta
CREATE TABLE [dbo].[DetaOVenta] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [Nroped] INT NULL,
    [Nro_Ocom] INT NULL,
    [Codint] INT NULL,
    [CodiExt] NVARCHAR(24) NULL,
    [Descrip] NVARCHAR(64) NULL,
    [CanPed] DECIMAL(18,4) NULL,
    [CantEnt] DECIMAL(18,4) NULL,
    [FechaModif] DATETIME2 NULL,
    [Accion] NVARCHAR(24) NULL,
    [NCtePedido] INT NULL,
    [MARBORRA] SMALLINT NULL,
    [CodiEmpr] INT NULL DEFAULT 0,
    [AddRecord] DATETIME2 NOT NULL DEFAULT GETDATE(),
    [LastUpdate] DATETIME2 NOT NULL DEFAULT GETDATE(),
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,
    CONSTRAINT PK_DetaOVenta_Id PRIMARY KEY (Id)
);

-- ÍNDICES PARA MEJORAR RENDIMIENTO
-- Índice para filtrar por número de pedido
CREATE NONCLUSTERED INDEX IX_DetaOVenta_Nroped ON [dbo].[DetaOVenta] ([Nroped]);

-- Índice para filtrar por número de orden de compra
CREATE NONCLUSTERED INDEX IX_DetaOVenta_NroOcom ON [dbo].[DetaOVenta] ([Nro_Ocom]);

-- Índice para filtrar por código interno
CREATE NONCLUSTERED INDEX IX_DetaOVenta_Codint ON [dbo].[DetaOVenta] ([Codint]);

-- Índice para filtrar registros no borrados
CREATE NONCLUSTERED INDEX IX_DetaOVenta_MARBORRA ON [dbo].[DetaOVenta] ([MARBORRA]);

-- Índice compuesto para consultas comunes
CREATE NONCLUSTERED INDEX IX_DetaOVenta_Nroped_MARBORRA ON [dbo].[DetaOVenta] ([Nroped], [MARBORRA]);

-- NOTAS IMPORTANTES:
-- 1. La tabla se crea con un campo Id IDENTITY como clave primaria
-- 2. Los campos de fecha usan DATETIME2 para mayor precisión
-- 3. Los campos de texto usan NVARCHAR para soportar caracteres Unicode
-- 4. CodiEmpr tiene valor por defecto 0 según las reglas del sistema
-- 5. Los campos de auditoría (AddRecord, LastUpdate, CodiEmprNet, NumUsuar) se agregan automáticamente
-- 6. Los índices mejoran significativamente el rendimiento de las consultas
*/
