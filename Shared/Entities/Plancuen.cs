using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa el plan de cuentas contables en el sistema.
    /// Mapea a la tabla PLANCUEN en la base de datos.
    /// </summary>
    [Table("PLANCUEN")]
    public class Plancuen
    {
        /// <summary>
        /// Identificador único de la cuenta (campo IDENTITY)
        /// </summary>
        [Key]
        [Column("Id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Código externo de la cuenta
        /// </summary>
        [Column("CodExCue")]
        [MaxLength(16)]
        public string? CodExCue { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Denominación de la cuenta
        /// </summary>
        [Column("Denocue")]
        [MaxLength(120)]
        public string? Denocue { get; set; }

        /// <summary>
        /// Código interno de la cuenta
        /// </summary>
        [Column("CodInCue")]
        public int? CodInCue { get; set; }

        /// <summary>
        /// Estado de la cuenta
        /// </summary>
        [Column("Statcue")]
        public short? Statcue { get; set; } = 1;

        /// <summary>
        /// Nivel de agrupación
        /// </summary>
        [Column("NivAgru")]
        public short? NivAgru { get; set; }

        /// <summary>
        /// Tipo de cuenta
        /// </summary>
        [Column("TipoCue")]
        [MaxLength(8)]
        public string? TipoCue { get; set; }

        /// <summary>
        /// Número de orden del item
        /// </summary>
        [Column("NuOrdIt")]
        public short? NuOrdIt { get; set; }

        /// <summary>
        /// Posición de impuesto
        /// </summary>
        [Column("PosImpu")]
        public short? PosImpu { get; set; }

        /// <summary>
        /// Observaciones de la cuenta
        /// </summary>
        [Column("Observa")]
        public string? Observa { get; set; }

        /// <summary>
        /// Saldo inicial
        /// </summary>
        [Column("Salini", TypeName = "decimal(18,4)")]
        public decimal? Salini { get; set; }

        /// <summary>
        /// Suma del debe
        /// </summary>
        [Column("Sudebe", TypeName = "decimal(18,4)")]
        public decimal? Sudebe { get; set; }

        /// <summary>
        /// Suma del haber
        /// </summary>
        [Column("Suhabe", TypeName = "decimal(18,4)")]
        public decimal? Suhabe { get; set; }

        /// <summary>
        /// Resultado
        /// </summary>
        [Column("Result", TypeName = "decimal(18,4)")]
        public decimal? Result { get; set; }

        /// <summary>
        /// Saldo final
        /// </summary>
        [Column("Salfin", TypeName = "decimal(18,4)")]
        public decimal? Salfin { get; set; }

        /// <summary>
        /// Indicador de ajuste
        /// </summary>
        [Column("IndiAjus")]
        public short? IndiAjus { get; set; }

        /// <summary>
        /// Moneda extranjera
        /// </summary>
        [Column("MoneExtr")]
        public short? MoneExtr { get; set; }

        /// <summary>
        /// Código de empresa .NET
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

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
        /// Identificador único global de la fila
        /// </summary>
        [Column("rowguid")]
        public Guid Rowguid { get; set; } = Guid.NewGuid();
    }
}

/*
-- Instrucciones SQL para modificar la tabla PLANCUEN en SQL Server Management Studio:

-- =======================================================================================
-- SECCIÓN 1: MODIFICAR CAMPOS EXISTENTES
-- =======================================================================================

-- Cambiar CodiEmpr de SMALLINT a INT
ALTER TABLE [dbo].[PLANCUEN] ALTER COLUMN [CodiEmpr] INT NULL;

-- Cambiar CodInCue de SMALLINT a INT  
ALTER TABLE [dbo].[PLANCUEN] ALTER COLUMN [CodInCue] INT NULL;

-- Cambiar campos MONEY a DECIMAL(18,4)
ALTER TABLE [dbo].[PLANCUEN] ALTER COLUMN [Salini] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[PLANCUEN] ALTER COLUMN [Sudebe] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[PLANCUEN] ALTER COLUMN [Suhabe] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[PLANCUEN] ALTER COLUMN [Result] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[PLANCUEN] ALTER COLUMN [Salfin] DECIMAL(18,4) NULL;

-- =======================================================================================
-- SECCIÓN 2: AGREGAR NUEVOS CAMPOS
-- =======================================================================================

-- Verificar y agregar campo ID IDENTITY si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = 'Id')
    ALTER TABLE [dbo].[PLANCUEN] ADD [Id] INT IDENTITY(1,1) NOT NULL;

-- Verificar y agregar CodiEmprNet si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = 'CodiEmprNet')
    ALTER TABLE [dbo].[PLANCUEN] ADD [CodiEmprNet] INT NULL;

-- Verificar y agregar NumUsuar si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = 'NumUsuar')
    ALTER TABLE [dbo].[PLANCUEN] ADD [NumUsuar] INT NULL;

-- Verificar y agregar AddRecord si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = 'AddRecord')
    ALTER TABLE [dbo].[PLANCUEN] ADD [AddRecord] DATETIME2 NULL;

-- Verificar y agregar LastUpdate si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = 'LastUpdate')
    ALTER TABLE [dbo].[PLANCUEN] ADD [LastUpdate] DATETIME2 NULL;

-- =======================================================================================
-- SECCIÓN 3: CREAR/ACTUALIZAR ÍNDICES Y CLAVE PRIMARIA
-- =======================================================================================

-- Crear clave primaria en el campo Id si no existe
IF NOT EXISTS (SELECT * FROM sys.key_constraints WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = N'PK_PLANCUEN')
    ALTER TABLE [dbo].[PLANCUEN] ADD CONSTRAINT [PK_PLANCUEN] PRIMARY KEY CLUSTERED ([Id] ASC);

-- Crear índice en rowguid si no existe
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = N'IX_PLANCUEN_rowguid')
    CREATE UNIQUE INDEX [IX_PLANCUEN_rowguid] ON [dbo].[PLANCUEN] ([rowguid] ASC);

-- Crear índice en código externo para búsquedas
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = N'IX_PLANCUEN_CodExCue')
    CREATE INDEX [IX_PLANCUEN_CodExCue] ON [dbo].[PLANCUEN] ([CodExCue] ASC);

-- Recrear índice en código interno (ahora INT en lugar de SMALLINT)
IF EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = N'IX_PLANCUEN_CodInCue')
    DROP INDEX [IX_PLANCUEN_CodInCue] ON [dbo].[PLANCUEN];
CREATE INDEX [IX_PLANCUEN_CodInCue] ON [dbo].[PLANCUEN] ([CodInCue] ASC);

-- Recrear índice compuesto para empresa y estado (CodiEmpr ahora es INT)
IF EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = N'IX_PLANCUEN_Empresa_Estado')
    DROP INDEX [IX_PLANCUEN_Empresa_Estado] ON [dbo].[PLANCUEN];
CREATE INDEX [IX_PLANCUEN_Empresa_Estado] ON [dbo].[PLANCUEN] ([CodiEmpr] ASC, [Statcue] ASC);

-- Crear índice en CodiEmprNet para búsquedas
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = N'IX_PLANCUEN_CodiEmprNet')
    CREATE INDEX [IX_PLANCUEN_CodiEmprNet] ON [dbo].[PLANCUEN] ([CodiEmprNet] ASC);

-- Crear índice en NumUsuar para auditoría
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[PLANCUEN]') AND name = N'IX_PLANCUEN_NumUsuar')
    CREATE INDEX [IX_PLANCUEN_NumUsuar] ON [dbo].[PLANCUEN] ([NumUsuar] ASC);

-- =======================================================================================
-- ESTRUCTURA FINAL DE LA TABLA (REFERENCIA)
-- =======================================================================================

-- CREATE TABLE [dbo].[PLANCUEN](
--     [Id] [int] IDENTITY(1,1) NOT NULL,        -- NUEVO CAMPO - Clave primaria
--     [CodExCue] [nvarchar](16) NULL,
--     [CodiEmpr] [int] NULL,                    -- Cambiado de smallint a int
--     [Denocue] [nvarchar](120) NULL,
--     [CodInCue] [int] NULL,                    -- Cambiado de smallint a int
--     [Statcue] [smallint] NULL,
--     [NivAgru] [smallint] NULL,
--     [TipoCue] [nvarchar](8) NULL,
--     [NuOrdIt] [smallint] NULL,
--     [PosImpu] [smallint] NULL,
--     [Observa] [ntext] NULL,
--     [Salini] [decimal](18,4) NULL,            -- Cambiado de money a decimal(18,4)
--     [Sudebe] [decimal](18,4) NULL,            -- Cambiado de money a decimal(18,4)
--     [Suhabe] [decimal](18,4) NULL,            -- Cambiado de money a decimal(18,4)
--     [Result] [decimal](18,4) NULL,            -- Cambiado de money a decimal(18,4)
--     [Salfin] [decimal](18,4) NULL,            -- Cambiado de money a decimal(18,4)
--     [IndiAjus] [smallint] NULL,
--     [MoneExtr] [smallint] NULL,
--     [CodiEmprNet] [int] NULL,                 -- NUEVO CAMPO
--     [NumUsuar] [int] NULL,                    -- NUEVO CAMPO
--     [AddRecord] [datetime2] NULL,             -- NUEVO CAMPO
--     [LastUpdate] [datetime2] NULL,            -- NUEVO CAMPO
--     [rowguid] [uniqueidentifier] ROWGUIDCOL NOT NULL,
--     CONSTRAINT [PK_PLANCUEN] PRIMARY KEY CLUSTERED ([Id] ASC)
-- ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

*/
