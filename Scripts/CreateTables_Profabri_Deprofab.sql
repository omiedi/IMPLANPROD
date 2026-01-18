/*
====================================================================
SCRIPT DE CREACIÓN DE TABLAS: PROFABRI Y DEPROFAB
Sistema: IMPLANPROD - Módulo de Planeamiento de Capacidad
Fecha: 2024-11-24
====================================================================
*/

-- TABLA: PROFABRI (Encabezados de Programas de Fabricación)
-- Almacena los programas de fabricación con su información básica
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PROFABRI]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[PROFABRI] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [NumeroPrograma] INT NOT NULL,
        [Descripcion] NVARCHAR(48) NOT NULL,
        [FechaCreacion] DATETIME NULL,
        [CodiEmpr] INT NULL,
        [Status] INT NULL DEFAULT 0,
        -- Campos de auditoría
        [AddRecord] DATETIME NULL,
        [LastUpdate] DATETIME NULL,
        [CodiEmprNet] INT NULL DEFAULT 0,
        [NumUsuar] INT NULL,
        
        CONSTRAINT [PK_PROFABRI] PRIMARY KEY CLUSTERED ([Id] ASC),
        CONSTRAINT [UQ_PROFABRI_NumeroPrograma] UNIQUE ([NumeroPrograma])
    );

    PRINT 'Tabla PROFABRI creada exitosamente.';
END
ELSE
BEGIN
    PRINT 'La tabla PROFABRI ya existe.';
END
GO

-- Índices para PROFABRI
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_PROFABRI_NumeroPrograma' AND object_id = OBJECT_ID('PROFABRI'))
BEGIN
    CREATE INDEX IX_PROFABRI_NumeroPrograma ON PROFABRI(NumeroPrograma);
    PRINT 'Índice IX_PROFABRI_NumeroPrograma creado.';
END
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_PROFABRI_CodiEmpr_Status' AND object_id = OBJECT_ID('PROFABRI'))
BEGIN
    CREATE INDEX IX_PROFABRI_CodiEmpr_Status ON PROFABRI(CodiEmpr, Status);
    PRINT 'Índice IX_PROFABRI_CodiEmpr_Status creado.';
END
GO

-- TABLA: DEPROFAB (Detalles de Programas de Fabricación)
-- Almacena los productos y cantidades de cada programa
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DEPROFAB]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[DEPROFAB] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [NumeroPrograma] INT NOT NULL,
        [CodigoProducto] INT NOT NULL,
        [Descripcion] NVARCHAR(64) NULL,
        [Cantidad] DECIMAL(18,4) NOT NULL,
        [Observaciones] NVARCHAR(52) NULL,
        -- Campos de auditoría
        [AddRecord] DATETIME NULL,
        [LastUpdate] DATETIME NULL,
        [CodiEmprNet] INT NULL DEFAULT 0,
        [NumUsuar] INT NULL,
        
        CONSTRAINT [PK_DEPROFAB] PRIMARY KEY CLUSTERED ([Id] ASC),
        
        -- Foreign key a PROFABRI
        CONSTRAINT [FK_DEPROFAB_PROFABRI] 
            FOREIGN KEY ([NumeroPrograma]) 
            REFERENCES [PROFABRI]([NumeroPrograma])
            ON DELETE CASCADE
    );

    PRINT 'Tabla DEPROFAB creada exitosamente.';
END
ELSE
BEGIN
    PRINT 'La tabla DEPROFAB ya existe.';
END
GO

-- Índices para DEPROFAB
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_DEPROFAB_NumeroPrograma' AND object_id = OBJECT_ID('DEPROFAB'))
BEGIN
    CREATE INDEX IX_DEPROFAB_NumeroPrograma ON DEPROFAB(NumeroPrograma);
    PRINT 'Índice IX_DEPROFAB_NumeroPrograma creado.';
END
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_DEPROFAB_CodigoProducto' AND object_id = OBJECT_ID('DEPROFAB'))
BEGIN
    CREATE INDEX IX_DEPROFAB_CodigoProducto ON DEPROFAB(CodigoProducto);
    PRINT 'Índice IX_DEPROFAB_CodigoProducto creado.';
END
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_DEPROFAB_Programa_Producto' AND object_id = OBJECT_ID('DEPROFAB'))
BEGIN
    CREATE INDEX IX_DEPROFAB_Programa_Producto ON DEPROFAB(NumeroPrograma, CodigoProducto);
    PRINT 'Índice IX_DEPROFAB_Programa_Producto creado.';
END
GO

-- Verificación de las tablas creadas
SELECT 
    'PROFABRI' AS Tabla,
    COUNT(*) AS NumeroColumnas
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'PROFABRI'
UNION ALL
SELECT 
    'DEPROFAB' AS Tabla,
    COUNT(*) AS NumeroColumnas
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'DEPROFAB';
GO

PRINT '====================================================================';
PRINT 'Script finalizado. Tablas PROFABRI y DEPROFAB listas para usar.';
PRINT '====================================================================';
GO

/*
====================================================================
NOTAS DE USO:
====================================================================

1. PROFABRI almacena los encabezados de programas:
   - NumeroPrograma: Número único autoincremental del programa
   - Descripcion: Nombre descriptivo del programa
   - Status: 0=Borrador, 1=Activo, 8=Anulado
   - Campos de auditoría se llenan automáticamente por middleware

2. DEPROFAB almacena los detalles (productos):
   - Cada registro representa un producto en el programa
   - La FK con CASCADE DELETE elimina automáticamente los detalles
     cuando se elimina el programa padre
   - Campos de auditoría se llenan automáticamente por middleware

3. Para generar un nuevo número de programa:
   SELECT ISNULL(MAX(NumeroPrograma), 0) + 1 FROM PROFABRI

4. Campos de auditoría:
   - AddRecord: Fecha de creación (se llena automáticamente)
   - LastUpdate: Fecha de última modificación (se actualiza automáticamente)
   - CodiEmprNet: Código de empresa del usuario (.NET)
   - NumUsuar: ID del usuario que realizó la acción

5. Ejemplo de inserción:
   
   -- Crear programa (los campos de auditoría se llenan automáticamente)
   INSERT INTO PROFABRI (NumeroPrograma, Descripcion, FechaCreacion, CodiEmpr, Status)
   VALUES (1, 'Programa Mayo 2024', GETDATE(), 1, 1);
   
   -- Agregar productos al programa
   INSERT INTO DEPROFAB (NumeroPrograma, CodigoProducto, Descripcion, Cantidad)
   VALUES 
       (1, 1001, 'Producto A', 100.00),
       (1, 1002, 'Producto B', 50.00);

====================================================================
*/
