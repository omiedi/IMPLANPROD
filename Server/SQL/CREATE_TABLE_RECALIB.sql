-- =====================================================
-- SCRIPT PARA CREAR LA TABLA RECALIB
-- Registros de calibración de instrumentos de medición
-- Migrado desde: Legacy_VB6/ArchivosDat/recalib.dat
-- =====================================================

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
GO

-- Índice para búsquedas por instrumento
CREATE INDEX IX_RECALIB_CodIntMedIdent ON [dbo].[RECALIB] ([CodIntMedIdent]);
GO

-- =====================================================
-- DESCRIPCIÓN DE CAMPOS
-- =====================================================
-- Id: Clave primaria autogenerada (IDENTITY)
-- RegCalibr_Nro: Número de registro de calibración (del archivo .dat)
-- FeCalibracion: Fecha en que se realizó la calibración
-- FeProxCalibr: Fecha programada para la próxima calibración
-- IC_NRO: Número de instrumento
-- PatronUtilizado: Patrón utilizado para la calibración
-- RequirioCalib: Indicador si requirió calibración
-- RequirioAjuste: Indicador si requirió ajuste
-- OBSERVACIONES: Observaciones del registro de calibración
-- RESPONSABLE: Responsable de la calibración
-- CodIntMedIdent: Código interno del instrumento (relación con INSTRUME.CodIntMedIdent)
-- Status: Estado del registro (0 = Activo, -1 = Eliminado)
-- AddRecord: Fecha de creación del registro
-- LastUpdate: Fecha de última actualización
-- CodiEmprNet: Código de empresa (siempre 0)
-- NumUsuar: Número de usuario que creó/modificó el registro
-- =====================================================
