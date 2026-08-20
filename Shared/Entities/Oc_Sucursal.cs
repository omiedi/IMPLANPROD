using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad OC_SUCURSAL - Tabla de órdenes de compra de sucursal
    /// </summary>
    [Table("OC_SUCURSAL")]
    public class Oc_Sucursal : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador interno autoincremental (IDENTITY) y clave primaria.
        /// Se agrega para usar una PK numérica por rendimiento y mantener los campos originales como datos de negocio.
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Código de sucursal
        /// </summary>
        [Column("SUCURSAL")]
        public short? Sucursal { get; set; }

        /// <summary>
        /// Fecha de emisión
        /// </summary>
        [Column("FECHEMISI")]
        public DateTime? Fechemisi { get; set; }

        /// <summary>
        /// Código interno del producto
        /// </summary>
        [Column("Codint")]
        public int? Codint { get; set; }

        /// <summary>
        /// Número de pedido de sucursal
        /// </summary>
        [Column("NROPED_SUC")]
        public int? NropedSuc { get; set; }

        /// <summary>
        /// Número de cliente del pedido
        /// </summary>
        [Column("NCLIEN_PEDIDO")]
        public int? NclienPedido { get; set; }

        /// <summary>
        /// Razón social del cliente de sucursal
        /// </summary>
        [Column("RAZ_CTE_SUC")]
        [MaxLength(32)]
        public string? RazCteSuc { get; set; }

        /// <summary>
        /// Usuario que creó el registro
        /// </summary>
        [Column("USUARIO")]
        public int? Usuario { get; set; }

        /// <summary>
        /// Orden de compra de sucursal
        /// </summary>
        [Column("O_COMPRA_SUC")]
        public int? OCompraSuc { get; set; }

        /// <summary>
        /// Cantidad solicitada
        /// </summary>
        [Column("CANTIDAD", TypeName = "decimal(18,4)")]
        public decimal? Cantidad { get; set; }

        /// <summary>
        /// Número de pedido generado en Avellaneda
        /// </summary>
        [Column("NPED_GENERADO_EN_AVELLANEDA")]
        public int? NpedGeneradoEnAvellaneda { get; set; }

        /// <summary>
        /// Versión de la compra de sucursal
        /// </summary>
        [Column("VERSION_COMPRA_SUC")]
        public short? VersionCompraSuc { get; set; }

        /// <summary>
        /// Marca de borrado
        /// </summary>
        [Column("MARBORRA")]
        public short? Marborra { get; set; }

        /// <summary>
        /// Estado del registro
        /// </summary>
        [Column("STATUS")]
        public short? Status { get; set; }

        /// <summary>
        /// Fecha de solicitud de compra
        /// </summary>
        [Column("FECHA_SOLICI_COMPRA")]
        public DateTime? FechaSoliciCompra { get; set; }

        /// <summary>
        /// Orden de compra del cliente
        /// </summary>
        [Column("OComCliente")]
        [MaxLength(16)]
        public string? OComCliente { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        [Column("COD_EXTE")]
        [MaxLength(24)]
        public string? CodExte { get; set; }

        /// <summary>
        /// Código de empresa
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
Ejecutar en SQL Server Management Studio para crear la tabla OC_SUCURSAL:

-- CREAR TABLA OC_SUCURSAL
CREATE TABLE [dbo].[OC_SUCURSAL] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [SUCURSAL] SMALLINT NULL,
    [FECHEMISI] DATETIME2 NULL,
    [Codint] INT NULL,
    [NROPED_SUC] INT NULL,
    [NCLIEN_PEDIDO] INT NULL,
    [RAZ_CTE_SUC] NVARCHAR(32) NULL,
    [USUARIO] INT NULL,
    [O_COMPRA_SUC] INT NULL,
    [CANTIDAD] DECIMAL(18,4) NULL,
    [NPED_GENERADO_EN_AVELLANEDA] INT NULL,
    [VERSION_COMPRA_SUC] SMALLINT NULL,
    [MARBORRA] SMALLINT NULL,
    [STATUS] SMALLINT NULL,
    [FECHA_SOLICI_COMPRA] DATETIME2 NULL,
    [OComCliente] NVARCHAR(16) NULL,
    [COD_EXTE] NVARCHAR(24) NULL,
    [CodiEmpr] INT NULL DEFAULT 0,
    [AddRecord] DATETIME2 NOT NULL DEFAULT GETDATE(),
    [LastUpdate] DATETIME2 NOT NULL DEFAULT GETDATE(),
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,
    CONSTRAINT PK_OC_SUCURSAL_Id PRIMARY KEY (Id)
);

-- ÍNDICES PARA MEJORAR RENDIMIENTO
-- Estos índices optimizan las consultas de la página OC Sucursal

-- Índice para filtrar registros no borrados (usado en WHERE Marborra != -1)
CREATE NONCLUSTERED INDEX IX_OC_SUCURSAL_MARBORRA ON [dbo].[OC_SUCURSAL] ([MARBORRA]);

-- Índice para JOIN con tabla de clientes (usado en JOIN ON Sucursal = Nume_Cli)
CREATE NONCLUSTERED INDEX IX_OC_SUCURSAL_SUCURSAL ON [dbo].[OC_SUCURSAL] ([SUCURSAL]);

-- Índice compuesto para ordenamiento (usado en ORDER BY VersionCompraSuc, OCompraSuc DESC)
CREATE NONCLUSTERED INDEX IX_OC_SUCURSAL_ORDENAMIENTO ON [dbo].[OC_SUCURSAL] ([VERSION_COMPRA_SUC], [O_COMPRA_SUC]);

-- Índice para filtrar por estado de pedido (usado en WHERE NpedGeneradoEnAvellaneda > 0 o <= 0)
CREATE NONCLUSTERED INDEX IX_OC_SUCURSAL_PEDIDO_GENERADO ON [dbo].[OC_SUCURSAL] ([NPED_GENERADO_EN_AVELLANEDA]);

-- Índice compuesto que cubre el filtro y ordenamiento más común
CREATE NONCLUSTERED INDEX IX_OC_SUCURSAL_MARBORRA_ORDEN ON [dbo].[OC_SUCURSAL] ([MARBORRA], [VERSION_COMPRA_SUC], [O_COMPRA_SUC]);

-- NOTAS IMPORTANTES:
-- 1. La tabla se crea con un campo Id IDENTITY como clave primaria
-- 2. Los campos de fecha usan DATETIME2 para mayor precisión
-- 3. Los campos de texto usan NVARCHAR para soportar caracteres Unicode
-- 4. CodiEmpr tiene valor por defecto 0 según las reglas del sistema
-- 5. Los campos de auditoría (AddRecord, LastUpdate, CodiEmprNet, NumUsuar) se agregan automáticamente
-- 6. Los índices mejoran significativamente el rendimiento de las consultas paginadas
*/
