using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa el detalle de IVA de órdenes de compra
    /// Equivalente a la tabla DETAIVA del VB6
    /// Implementa IAuditableEntity para auditoría automática
    /// </summary>
    [Table("DETAIVA")]
    public class Detaiva : IAuditableEntity
    {
        /// <summary>
        /// ID único del detalle de IVA (campo identity)
        /// </summary>
        [Key]
        [Column("DETAIVA_ID")]
        public int DETAIVA_ID { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Número de orden de compra
        /// </summary>
        [Column("nume_ocom")]
        public int? NumeOcom { get; set; }

        /// <summary>
        /// Número de orden de IVA (0-5)
        /// </summary>
        [Column("NORIT")]
        public int? Norit { get; set; }

        /// <summary>
        /// Tasa de IVA aplicada
        /// </summary>
        [Column("Tasa_Iva", TypeName = "decimal(18,4)")]
        public decimal? TasaIva { get; set; }

        /// <summary>
        /// Base imponible para este porcentaje de IVA
        /// </summary>
        [Column("Base_Impo", TypeName = "decimal(18,4)")]
        public decimal? BaseImpo { get; set; }

        /// <summary>
        /// Importe de IVA calculado
        /// </summary>
        [Column("Impo_Iva", TypeName = "decimal(18,4)")]
        public decimal? ImpoIva { get; set; }

        /// <summary>
        /// GUID único del registro
        /// </summary>
        [Column("rowguid")]
        public Guid Rowguid { get; set; } = Guid.NewGuid();

        /// <summary>
        /// Código de empresa del usuario que creó/modificó el registro
        /// Se asigna automáticamente por el AuditInterceptor
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// ID del usuario que creó/modificó el registro
        /// Se asigna automáticamente por el AuditInterceptor
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

        /// <summary>
        /// Fecha de creación del registro
        /// Se asigna automáticamente por el AuditInterceptor
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se asigna automáticamente por el AuditInterceptor
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }
    }
}

/*
-- =====================================================
-- INSTRUCCIONES SQL PARA CREAR LA TABLA DETAIVA
-- =====================================================

-- CREAR TABLA DETAIVA SI NO EXISTE
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='DETAIVA' AND xtype='U')
BEGIN
    CREATE TABLE [dbo].[DETAIVA] (
        [DETAIVA_ID] INT IDENTITY(1,1) NOT NULL,
        [CodiEmpr] INT NULL,
        [nume_ocom] INT NULL,
        [NORIT] INT NULL,
        [Tasa_Iva] DECIMAL(18,4) NULL,
        [Base_Impo] DECIMAL(18,4) NULL,
        [Impo_Iva] DECIMAL(18,4) NULL,
        [rowguid] UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID(),
        [CodiEmprNet] INT NULL,
        [NumUsuar] INT NULL,
        [AddRecord] DATETIME2 NOT NULL DEFAULT GETDATE(),
        [LastUpdate] DATETIME2 NOT NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_DETAIVA] PRIMARY KEY ([DETAIVA_ID])

    );
    
    -- Crear índices para optimizar consultas
    CREATE INDEX [IX_DETAIVA_NumeOcom] ON [dbo].[DETAIVA] ([nume_ocom]);
    CREATE INDEX [IX_DETAIVA_CodiEmpr] ON [dbo].[DETAIVA] ([CodiEmpr]);
    CREATE INDEX [IX_DETAIVA_CodiEmprNet] ON [dbo].[DETAIVA] ([CodiEmprNet]);
    CREATE INDEX [IX_DETAIVA_NumUsuar] ON [dbo].[DETAIVA] ([NumUsuar]);
END

-- AGREGAR CAMPOS DE AUDITORÍA SI LA TABLA YA EXISTE
IF EXISTS (SELECT * FROM sysobjects WHERE name='DETAIVA' AND xtype='U')
BEGIN
    -- Agregar CodiEmprNet si no existe
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('DETAIVA') AND name = 'CodiEmprNet')
        ALTER TABLE [dbo].[DETAIVA] ADD [CodiEmprNet] INT NULL;
    
    -- Agregar NumUsuar si no existe
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('DETAIVA') AND name = 'NumUsuar')
        ALTER TABLE [dbo].[DETAIVA] ADD [NumUsuar] INT NULL;
    
    -- Modificar AddRecord para que no sea nullable si es necesario
    IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('DETAIVA') AND name = 'AddRecord' AND is_nullable = 1)
        ALTER TABLE [dbo].[DETAIVA] ALTER COLUMN [AddRecord] DATETIME2 NOT NULL;
    
    -- Modificar LastUpdate para que no sea nullable si es necesario
    IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('DETAIVA') AND name = 'LastUpdate' AND is_nullable = 1)
        ALTER TABLE [dbo].[DETAIVA] ALTER COLUMN [LastUpdate] DATETIME2 NOT NULL;
END

-- NOTA: Ejecutar este comando en SQL Server Management Studio
-- para crear la tabla DETAIVA en la base de datos.
*/
