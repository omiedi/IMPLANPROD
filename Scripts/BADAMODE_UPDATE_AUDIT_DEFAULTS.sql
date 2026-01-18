-- =============================================
-- Script: Actualización de Campos de Auditoría por Defecto
-- Tabla: BADAMODE
-- Descripción: Aplica valores por defecto a registros ya migrados
-- Fecha: 2026-01-08
-- =============================================

-- =============================================
-- ACTUALIZACIÓN DE CAMPOS DE AUDITORÍA
-- =============================================
-- Este script actualiza los registros existentes en BADAMODE
-- que no tienen valores en los campos de auditoría

DECLARE @FechaActual DATETIME = GETDATE()
DECLARE @NumUsuarDefecto INT = 10
DECLARE @CodiEmprNetDefecto INT = 101
DECLARE @RegistrosActualizados INT = 0

PRINT '============================================='
PRINT 'Iniciando actualización de campos de auditoría'
PRINT 'Fecha: ' + CONVERT(VARCHAR(20), @FechaActual, 120)
PRINT '============================================='
PRINT ''

-- Verificar que la tabla existe
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND type in (N'U'))
BEGIN
    PRINT 'ERROR: La tabla BADAMODE no existe'
    PRINT 'Por favor ejecute primero el script BADAMODE_CREATE_AND_MODIFY.sql'
    RETURN
END

-- Verificar que los campos de auditoría existen
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE]') AND name = 'AddRecord')
BEGIN
    PRINT 'ERROR: El campo AddRecord no existe'
    PRINT 'Por favor ejecute primero el script BADAMODE_CREATE_AND_MODIFY.sql'
    RETURN
END

PRINT 'Tabla BADAMODE encontrada'
PRINT 'Campos de auditoría verificados'
PRINT ''

-- =============================================
-- ACTUALIZAR REGISTROS EXISTENTES
-- =============================================

BEGIN TRANSACTION

BEGIN TRY
    -- Actualizar registros que no tienen valores de auditoría
    UPDATE [dbo].[BADAMODE]
    SET 
        [AddRecord] = ISNULL([AddRecord], @FechaActual),
        [LastUpdate] = ISNULL([LastUpdate], @FechaActual),
        [NumUsuar] = ISNULL([NumUsuar], @NumUsuarDefecto),
        [CodiEmprNet] = ISNULL([CodiEmprNet], @CodiEmprNetDefecto)
    WHERE 
        [AddRecord] IS NULL 
        OR [LastUpdate] IS NULL 
        OR [NumUsuar] IS NULL 
        OR [CodiEmprNet] IS NULL

    SET @RegistrosActualizados = @@ROWCOUNT

    COMMIT TRANSACTION

    PRINT '============================================='
    PRINT 'Actualización completada exitosamente'
    PRINT '============================================='
    PRINT ''
    PRINT 'Registros actualizados: ' + CAST(@RegistrosActualizados AS VARCHAR(10))
    PRINT ''
    PRINT 'Valores aplicados:'
    PRINT '  - AddRecord: ' + CONVERT(VARCHAR(20), @FechaActual, 120)
    PRINT '  - LastUpdate: ' + CONVERT(VARCHAR(20), @FechaActual, 120)
    PRINT '  - NumUsuar: ' + CAST(@NumUsuarDefecto AS VARCHAR(10))
    PRINT '  - CodiEmprNet: ' + CAST(@CodiEmprNetDefecto AS VARCHAR(10))
    PRINT ''

    -- Mostrar resumen de registros actualizados
    SELECT 
        COUNT(*) AS 'Total Registros',
        COUNT(CASE WHEN AddRecord = @FechaActual THEN 1 END) AS 'Con AddRecord Actualizado',
        COUNT(CASE WHEN LastUpdate = @FechaActual THEN 1 END) AS 'Con LastUpdate Actualizado',
        COUNT(CASE WHEN NumUsuar = @NumUsuarDefecto THEN 1 END) AS 'Con NumUsuar = 10',
        COUNT(CASE WHEN CodiEmprNet = @CodiEmprNetDefecto THEN 1 END) AS 'Con CodiEmprNet = 101'
    FROM [dbo].[BADAMODE]

    PRINT ''
    PRINT '============================================='
    PRINT 'Script completado exitosamente'
    PRINT '============================================='

END TRY
BEGIN CATCH
    ROLLBACK TRANSACTION

    PRINT ''
    PRINT '============================================='
    PRINT 'ERROR: La actualización falló'
    PRINT '============================================='
    PRINT 'Mensaje de error: ' + ERROR_MESSAGE()
    PRINT 'Línea: ' + CAST(ERROR_LINE() AS VARCHAR(10))
    PRINT 'Procedimiento: ' + ISNULL(ERROR_PROCEDURE(), 'Script principal')
    PRINT ''
    PRINT 'La transacción ha sido revertida'
    PRINT 'No se realizaron cambios en la base de datos'
    PRINT '============================================='
END CATCH
GO

-- =============================================
-- VERIFICACIÓN FINAL
-- =============================================

-- Mostrar algunos registros de ejemplo
PRINT ''
PRINT 'Muestra de registros actualizados (primeros 10):'
PRINT ''

SELECT TOP 10
    Id,
    Cod_Mod,
    Descripcion,
    AddRecord,
    LastUpdate,
    NumUsuar,
    CodiEmprNet
FROM [dbo].[BADAMODE]
ORDER BY Id

GO
