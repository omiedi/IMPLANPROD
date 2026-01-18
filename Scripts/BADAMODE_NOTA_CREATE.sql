-- =============================================
-- Script: Creación de Tabla BADAMODE_NOTA
-- Descripción: Notas y observaciones de moldes
-- Fecha: 2026-01-08
-- =============================================

-- =============================================
-- PARTE 1: CREACIÓN DE LA TABLA
-- =============================================

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE_NOTA]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[BADAMODE_NOTA](
        -- Clave primaria autoincremental
        [Id] INT IDENTITY(1,1) NOT NULL,
        
        -- Campos de datos
        [Cod_Mod] NVARCHAR(24) NULL,
        [Observaciones] NVARCHAR(MAX) NULL,
        
        -- Campos de auditoría
        [AddRecord] DATETIME NULL,
        [LastUpdate] DATETIME NULL,
        [CodiEmprNet] INT NULL,
        [NumUsuar] INT NULL,
        
        -- Constraint de clave primaria
        CONSTRAINT [PK_BADAMODE_NOTA] PRIMARY KEY CLUSTERED ([Id] ASC)
    ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    
    PRINT 'Tabla BADAMODE_NOTA creada exitosamente'
END
ELSE
BEGIN
    PRINT 'La tabla BADAMODE_NOTA ya existe'
END
GO

-- =============================================
-- PARTE 2: CREAR ÍNDICES PARA OPTIMIZACIÓN
-- =============================================

-- Índice en Cod_Mod para búsquedas rápidas por molde
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE_NOTA]') AND name = 'IX_BADAMODE_NOTA_Cod_Mod')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BADAMODE_NOTA_Cod_Mod]
    ON [dbo].[BADAMODE_NOTA] ([Cod_Mod])
    
    PRINT 'Índice IX_BADAMODE_NOTA_Cod_Mod creado'
END
ELSE
BEGIN
    PRINT 'El índice IX_BADAMODE_NOTA_Cod_Mod ya existe'
END
GO

-- Índice en AddRecord para ordenamiento por fecha de creación
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE_NOTA]') AND name = 'IX_BADAMODE_NOTA_AddRecord')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BADAMODE_NOTA_AddRecord]
    ON [dbo].[BADAMODE_NOTA] ([AddRecord] DESC)
    
    PRINT 'Índice IX_BADAMODE_NOTA_AddRecord creado'
END
ELSE
BEGIN
    PRINT 'El índice IX_BADAMODE_NOTA_AddRecord ya existe'
END
GO

-- Índice compuesto para búsquedas por molde y fecha
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE_NOTA]') AND name = 'IX_BADAMODE_NOTA_Cod_Mod_AddRecord')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BADAMODE_NOTA_Cod_Mod_AddRecord]
    ON [dbo].[BADAMODE_NOTA] ([Cod_Mod], [AddRecord] DESC)
    
    PRINT 'Índice IX_BADAMODE_NOTA_Cod_Mod_AddRecord creado'
END
ELSE
BEGIN
    PRINT 'El índice IX_BADAMODE_NOTA_Cod_Mod_AddRecord ya existe'
END
GO

-- =============================================
-- PARTE 3: VERIFICACIÓN FINAL
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
WHERE c.object_id = OBJECT_ID(N'[dbo].[BADAMODE_NOTA]')
ORDER BY c.column_id

PRINT ''
PRINT '============================================='
PRINT 'Script completado exitosamente'
PRINT 'Tabla BADAMODE_NOTA lista para usar con Entity Framework'
PRINT '============================================='
PRINT ''
PRINT 'Características de la tabla:'
PRINT '  - Id: Clave primaria IDENTITY autoincremental'
PRINT '  - Cod_Mod: Código del molde (relación con BADAMODE)'
PRINT '  - Observaciones: Campo de texto largo (VARCHAR MAX)'
PRINT '  - Campos de auditoría: AddRecord, LastUpdate, CodiEmprNet, NumUsuar'
PRINT '  - Índices optimizados para búsquedas por molde y fecha'
PRINT '============================================='
GO
