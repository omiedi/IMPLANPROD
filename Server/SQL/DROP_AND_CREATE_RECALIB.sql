-- =====================================================
-- SCRIPT PARA BORRAR Y RECREAR LA TABLA RECALIB
-- =====================================================

-- Borrar la tabla si existe
IF OBJECT_ID('dbo.RECALIB', 'U') IS NOT NULL
BEGIN
    DROP TABLE [dbo].[RECALIB];
    PRINT 'Tabla RECALIB eliminada correctamente.';
END
ELSE
BEGIN
    PRINT 'La tabla RECALIB no existe.';
END
GO

-- Crear la tabla principal
CREATE TABLE [dbo].[RECALIB](
    [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [RegCalibr_Nro] INT NULL,
    [FeCalibracion] DATETIME NULL,
    [FeProxCalibr] DATETIME NULL,
    [IC_NRO] VARCHAR(16) NULL,
    [PatronUtilizado] VARCHAR(16) NULL,
    [RequirioCalib] VARCHAR(16) NULL,
    [RequirioAjuste] VARCHAR(16) NULL,
    [OBSERVACIONES] VARCHAR(MAX) NULL,
    [RESPONSABLE] VARCHAR(32) NULL,
    [CodIntMedIdent] INT NULL,
    [Status] SMALLINT NOT NULL DEFAULT 0,
    [AddRecord] DATETIME NULL,
    [LastUpdate] DATETIME NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL
);
PRINT 'Tabla RECALIB creada correctamente.';
GO

-- Índice para búsquedas por instrumento
CREATE INDEX IX_RECALIB_CodIntMedIdent ON [dbo].[RECALIB] ([CodIntMedIdent]);
PRINT 'Índice creado correctamente.';
GO

PRINT 'Script completado exitosamente.';
