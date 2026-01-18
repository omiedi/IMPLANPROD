-- =====================================================
-- SCRIPT PARA AGREGAR CAMPOS DE AUDITORÍA A OCOMENCA
-- =====================================================
-- Este script agrega los campos de auditoría faltantes
-- que están definidos en la entidad pero no en la BD

USE [TU_BASE_DE_DATOS] -- Reemplaza con el nombre de tu base de datos
GO

-- Verificar si la tabla OCOMENCA existe
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='OCOMENCA' AND xtype='U')
BEGIN
    PRINT 'ERROR: La tabla OCOMENCA no existe.'
    RETURN
END

-- 1. Agregar campo CodiEmprNet si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('OCOMENCA') AND name = 'CodiEmprNet')
BEGIN
    ALTER TABLE [dbo].[OCOMENCA] ADD [CodiEmprNet] INT NULL;
    PRINT 'Campo CodiEmprNet agregado exitosamente.';
END
ELSE
BEGIN
    PRINT 'Campo CodiEmprNet ya existe.';
END

-- 2. Agregar campo NumUsuar si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('OCOMENCA') AND name = 'NumUsuar')
BEGIN
    ALTER TABLE [dbo].[OCOMENCA] ADD [NumUsuar] INT NULL;
    PRINT 'Campo NumUsuar agregado exitosamente.';
END
ELSE
BEGIN
    PRINT 'Campo NumUsuar ya existe.';
END

-- 3. Agregar campo AddRecord si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('OCOMENCA') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[OCOMENCA] ADD [AddRecord] DATETIME2 NULL DEFAULT GETDATE();
    PRINT 'Campo AddRecord agregado exitosamente.';
END
ELSE
BEGIN
    PRINT 'Campo AddRecord ya existe.';
END

-- 4. Agregar campo LastUpdate si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('OCOMENCA') AND name = 'LastUpdate')
BEGIN
    ALTER TABLE [dbo].[OCOMENCA] ADD [LastUpdate] DATETIME2 NULL DEFAULT GETDATE();
    PRINT 'Campo LastUpdate agregado exitosamente.';
END
ELSE
BEGIN
    PRINT 'Campo LastUpdate ya existe.';
END

-- 5. Verificar los campos agregados
SELECT 
    COLUMN_NAME,
    DATA_TYPE,
    IS_NULLABLE,
    COLUMN_DEFAULT
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'OCOMENCA' 
    AND COLUMN_NAME IN ('CodiEmprNet', 'NumUsuar', 'AddRecord', 'LastUpdate')
ORDER BY COLUMN_NAME;

PRINT 'Script de campos de auditoría completado.';
