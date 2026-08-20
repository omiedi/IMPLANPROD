using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad para la gestión de moldes (BADAMODE)
    /// Representa la información de moldes utilizados en la producción
    /// </summary>
    [Table("BADAMODE")]
    public class Badamode : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// ID único del registro (clave primaria)
        /// </summary>
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        /// <summary>
        /// Código del molde
        /// </summary>
        [Column("Cod_Mod")]
        [StringLength(16)]
        public string? Cod_Mod { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Descripción del molde
        /// </summary>
        [Column("Descripcion")]
        [StringLength(64)]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Ubicación del molde
        /// </summary>
        [Column("Ubic")]
        [StringLength(32)]
        public string? Ubic { get; set; }

        /// <summary>
        /// Cantidad de piezas
        /// </summary>
        [Column("CantPiezas", TypeName = "decimal(18,4)")]
        public decimal? CantPiezas { get; set; }

        /// <summary>
        /// Frecuencia de control
        /// </summary>
        [Column("Frecuen")]
        [StringLength(32)]
        public string? Frecuen { get; set; }

        /// <summary>
        /// Fecha del último control
        /// </summary>
        [Column("FechUlt_Con")]
        public DateTime? FechUlt_Con { get; set; }

        /// <summary>
        /// Fecha del próximo control
        /// </summary>
        [Column("FechPro_Con")]
        public DateTime? FechPro_Con { get; set; }

        /// <summary>
        /// Código de material
        /// </summary>
        [Column("Material")]
        public int? Material { get; set; }

        /// <summary>
        /// Descripción del material
        /// </summary>
        [Column("Desc_Mat")]
        [StringLength(32)]
        public string? Desc_Mat { get; set; }

        /// <summary>
        /// Código interno
        /// </summary>
        [Column("Cod_Inte")]
        public int? Cod_Inte { get; set; }

        ///// <summary>
        ///// Anotaciones del molde
        ///// </summary>
        //[Column("AnotaMod")]
        //public string? AnotaMod { get; set; }

        /// <summary>
        /// Observaciones del molde
        /// </summary>
        [Column("ObservaMod", TypeName = "ntext")]
        public string? ObservaMod { get; set; }

        /// <summary>
        /// Ruta de la foto 1
        /// </summary>
        [Column("Foto1")]
        [StringLength(128)]
        public string? Foto1 { get; set; }

        /// <summary>
        /// Ruta de la foto 2
        /// </summary>
        [Column("Foto2")]
        [StringLength(128)]
        public string? Foto2 { get; set; }

        /// <summary>
        /// Ruta de la foto 3
        /// </summary>
        [Column("Foto3")]
        [StringLength(128)]
        public string? Foto3 { get; set; }

        /// <summary>
        /// Ruta de la foto 4
        /// </summary>
        [Column("Foto4")]
        [StringLength(128)]
        public string? Foto4 { get; set; }

        /// <summary>
        /// Ruta de la foto 2 (duplicado)
        /// </summary>
        [Column("Foto 2")]
        [StringLength(128)]
        public string? Foto_2 { get; set; }

        /// <summary>
        /// Ruta de la foto 3 (duplicado)
        /// </summary>
        [Column("Foto 3")]
        [StringLength(128)]
        public string? Foto_3 { get; set; }

        /// <summary>
        /// Ruta de la foto 4 (duplicado)
        /// </summary>
        [Column("Foto 4")]
        [StringLength(128)]
        public string? Foto_4 { get; set; }

        ///// <summary>
        ///// GUID único del registro
        ///// </summary>
        //[Column("rowguid")]
        //[Required]
        //public Guid Rowguid { get; set; } = Guid.NewGuid();

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// Se establece automáticamente cuando se crea el molde
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se actualiza automáticamente cada vez que se modifica el molde
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

        // ===================================================================
        // CAMPO DE ESTADO
        // ===================================================================

        /// <summary>
        /// Estado del molde
        /// Valores posibles: 0 = Inactivo, 1 = Activo, etc.
        /// </summary>
        [Column("Status")]
        public short? Status { get; set; }
    }

}
/*
-- =============================================
-- Script: Creación y Modificación de Tabla BADAMODE
-- Descripción: Gestión de moldes
-- Fecha: 2026-01-08
-- =============================================

-- =============================================
-- PARTE 1: CREACIÓN DE LA TABLA ORIGINAL
-- =============================================
-- Esta es la estructura original de la tabla BADAMODE
-- según el esquema proporcionado

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[BADAMODE](
        [Cod_Mod] [nvarchar](16) NULL,
        [CodiEmpr] [smallint] NULL,
        [Descripcion] [nvarchar](64) NULL,
        [Ubic] [nvarchar](32) NULL,
        [CantPiezas] [float] NULL,
        [Frecuen] [nvarchar](32) NULL,
        [FechUlt_Con] [smalldatetime] NULL,
        [FechPro_Con] [smalldatetime] NULL,
        [Material] [smallint] NULL,
        [Desc_Mat] [nvarchar](32) NULL,
        [Cod_Inte] [smallint] NULL,
        [AnotaMod] [ntext] NULL,
        [ObservaMod] [ntext] NULL,
        [Foto1] [nvarchar](128) NULL,
        [Foto2] [nvarchar](128) NULL,
        [Foto3] [nvarchar](128) NULL,
        [Foto4] [nvarchar](128) NULL,
        [Foto 2] [nvarchar](128) NULL,
        [Foto 3] [nvarchar](128) NULL,
        [Foto 4] [nvarchar](128) NULL,
        [rowguid] [uniqueidentifier] NOT NULL DEFAULT NEWID()
    ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    
    PRINT 'Tabla BADAMODE creada exitosamente'
END
ELSE
BEGIN
    PRINT 'La tabla BADAMODE ya existe'
END
GO

-- =============================================
-- PARTE 2: MODIFICACIONES DE TIPOS DE DATOS
-- =============================================
-- Modificar tipos de datos según los requerimientos de Entity Framework

-- 2.1: Agregar campo ID IDENTITY si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'Id')
BEGIN
    -- Crear columna Id como IDENTITY
    ALTER TABLE [dbo].[BADAMODE]
    ADD [Id] INT IDENTITY(1,1) NOT NULL
    
    -- Agregar constraint de clave primaria
    ALTER TABLE [dbo].[BADAMODE]
    ADD CONSTRAINT PK_BADAMODE PRIMARY KEY CLUSTERED ([Id])
    
    PRINT 'Campo Id agregado como clave primaria IDENTITY'
END
ELSE
BEGIN
    PRINT 'El campo Id ya existe'
END
GO

-- 2.2: Modificar CodiEmpr de SMALLINT a INT
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'CodiEmpr' AND system_type_id = 52)
BEGIN
    ALTER TABLE [dbo].[BADAMODE]
    ALTER COLUMN [CodiEmpr] INT NULL
    
    PRINT 'Campo CodiEmpr modificado de SMALLINT a INT'
END
ELSE
BEGIN
    PRINT 'Campo CodiEmpr ya es INT o no existe'
END
GO

-- 2.3: Modificar Cod_Inte de SMALLINT a INT
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'Cod_Inte' AND system_type_id = 52)
BEGIN
    ALTER TABLE [dbo].[BADAMODE]
    ALTER COLUMN [Cod_Inte] INT NULL
    
    PRINT 'Campo Cod_Inte modificado de SMALLINT a INT'
END
ELSE
BEGIN
    PRINT 'Campo Cod_Inte ya es INT o no existe'
END
GO

-- 2.4: Modificar Material de SMALLINT a INT
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'Material' AND system_type_id = 52)
BEGIN
    ALTER TABLE [dbo].[BADAMODE]
    ALTER COLUMN [Material] INT NULL
    
    PRINT 'Campo Material modificado de SMALLINT a INT'
END
ELSE
BEGIN
    PRINT 'Campo Material ya es INT o no existe'
END
GO

-- 2.5: Modificar CantPiezas de FLOAT a DECIMAL(18,4)
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'CantPiezas' AND system_type_id = 62)
BEGIN
    ALTER TABLE [dbo].[BADAMODE]
    ALTER COLUMN [CantPiezas] DECIMAL(18,4) NULL
    
    PRINT 'Campo CantPiezas modificado de FLOAT a DECIMAL(18,4)'
END
ELSE
BEGIN
    PRINT 'Campo CantPiezas ya es DECIMAL o no existe'
END
GO

-- =============================================
-- PARTE 3: AGREGAR CAMPOS DE AUDITORÍA
-- =============================================
-- Campos para implementar ITimestampEntity y IAuditableEntity

-- 3.1: Agregar campo AddRecord (Fecha de creación)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[BADAMODE]
    ADD [AddRecord] DATETIME NULL
    
    PRINT 'Campo AddRecord agregado'
END
ELSE
BEGIN
    PRINT 'El campo AddRecord ya existe'
END
GO

-- 3.2: Agregar campo LastUpdate (Fecha de última actualización)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'LastUpdate')
BEGIN
    ALTER TABLE [dbo].[BADAMODE]
    ADD [LastUpdate] DATETIME NULL
    
    PRINT 'Campo LastUpdate agregado'
END
ELSE
BEGIN
    PRINT 'El campo LastUpdate ya existe'
END
GO

-- 3.3: Agregar campo CodiEmprNet (Código de empresa del usuario)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'CodiEmprNet')
BEGIN
    ALTER TABLE [dbo].[BADAMODE]
    ADD [CodiEmprNet] INT NULL
    
    PRINT 'Campo CodiEmprNet agregado'
END
ELSE
BEGIN
    PRINT 'El campo CodiEmprNet ya existe'
END
GO

-- 3.4: Agregar campo NumUsuar (Número de usuario)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'NumUsuar')
BEGIN
    ALTER TABLE [dbo].[BADAMODE]
    ADD [NumUsuar] INT NULL
    
    PRINT 'Campo NumUsuar agregado'
END
ELSE
BEGIN
    PRINT 'El campo NumUsuar ya existe'
END
GO

-- =============================================
-- PARTE 4: CREAR ÍNDICES PARA OPTIMIZACIÓN
-- =============================================

-- Índice en Cod_Mod para búsquedas rápidas
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'IX_BADAMODE_Cod_Mod')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BADAMODE_Cod_Mod]
    ON [dbo].[BADAMODE] ([Cod_Mod])
    
    PRINT 'Índice IX_BADAMODE_Cod_Mod creado'
END
ELSE
BEGIN
    PRINT 'El índice IX_BADAMODE_Cod_Mod ya existe'
END
GO

-- Índice en CodiEmpr para filtrado por empresa
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'IX_BADAMODE_CodiEmpr')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BADAMODE_CodiEmpr]
    ON [dbo].[BADAMODE] ([CodiEmpr])
    
    PRINT 'Índice IX_BADAMODE_CodiEmpr creado'
END
ELSE
BEGIN
    PRINT 'El índice IX_BADAMODE_CodiEmpr ya existe'
END
GO

-- Índice en FechPro_Con para control de próximos mantenimientos
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'IX_BADAMODE_FechPro_Con')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BADAMODE_FechPro_Con]
    ON [dbo].[BADAMODE] ([FechPro_Con])
    
    PRINT 'Índice IX_BADAMODE_FechPro_Con creado'
END
ELSE
BEGIN
    PRINT 'El índice IX_BADAMODE_FechPro_Con ya existe'
END
GO

-- =============================================
-- PARTE 5: VERIFICACIÓN FINAL
-- =============================================

-- Mostrar estructura final de la tabla
SELECT 
    c.name AS 'Columna',
    t.name AS 'Tipo',
    c.max_length AS 'Longitud',
    c.is_nullable AS 'Permite NULL',
    c.is_identity AS 'Es IDENTITY'
FROM sys.columns c
INNER JOIN sys.types t ON c.user_type_id = t.user_type_id
WHERE c.object_id = OBJECT_ID(N'[dbo].[BADAMODE]')
ORDER BY c.column_id

PRINT '============================================='
PRINT 'Script completado exitosamente'
PRINT 'Tabla BADAMODE lista para usar con Entity Framework'
PRINT '============================================='
GO

-- =============================================
-- AGREGAR CAMPO STATUS
-- =============================================
-- Agregar campo Status de tipo SMALLINT si no existe

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'Status')
BEGIN
    ALTER TABLE [dbo].[BADAMODE]
    ADD [Status] SMALLINT NULL
    
    PRINT 'Campo Status agregado exitosamente'
END
ELSE
BEGIN
    PRINT 'El campo Status ya existe'
END
GO

-- Crear índice en Status para filtrado rápido por estado
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'IX_BADAMODE_Status')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BADAMODE_Status]
    ON [dbo].[BADAMODE] ([Status])
    
    PRINT 'Índice IX_BADAMODE_Status creado'
END
ELSE
BEGIN
    PRINT 'El índice IX_BADAMODE_Status ya existe'
END
GO

PRINT '============================================='
PRINT 'Campo Status agregado correctamente'
PRINT 'Valores sugeridos: 0 = Inactivo, 1 = Activo'
PRINT '============================================='
GO
==================================================
********* PARA PONERLE VALOR A ESTOS CAMPOS *****************
  UPDATE BADAMODE
SET Status = 0;

UPDATE BADAMODE
SET LastUpdate = GETDATE(),
    AddRecord  = GETDATE();

	UPDATE BADAMODE
SET NumUsuar      = 10,
    [CodiEmprNet] = 101;
***********************   pasada

*/
/*  ESTO ES LPARA CARGAR LAS NOTAS EN LA TABLA BADAMODE_NOTA  
 *  --1) Borrar el contenido de la tabla destino
DELETE FROM arcon2.dbo.BADAMODE_NOTA;


--2) Insertar los datos desde la tabla original
INSERT INTO arcon2.dbo.BADAMODE_NOTA (COD_MOD, OBSERVACIONES)
SELECT
    COD_MOD,
    ANOTAMOD
FROM arcon.dbo.BADAMODE
WHERE ANOTAMOD IS NOT NULL
  AND CAST(ANOTAMOD AS NVARCHAR(MAX)) <> '';
*/