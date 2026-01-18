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
