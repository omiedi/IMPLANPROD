-- =============================================
-- Script: Optimización de performance para O.F. Limpieza
-- Fecha: 2024-12-16
-- Descripción: Agrega índice en ORDEFABR.IdSubOr para mejorar búsquedas por prefijo (DL-)
-- =============================================

USE [IMPLANPROD]
GO

-- Verificar si el índice ya existe
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_ORDEFABR_IdSubOr' AND object_id = OBJECT_ID('ORDEFABR'))
BEGIN
    PRINT 'Creando índice IX_ORDEFABR_IdSubOr...'
    
    CREATE NONCLUSTERED INDEX [IX_ORDEFABR_IdSubOr]
    ON [dbo].[ORDEFABR] ([IdSubOr])
    INCLUDE ([FechGen], [StatuOrf], [Cod_Exte], [Unidad], [FechEnSol], [FechLan], [FechCierr], [Descrip], [Cod_Inte], [CanProy], [CanApro], [Id_Orfa])
    GO
    
    PRINT 'Índice IX_ORDEFABR_IdSubOr creado exitosamente.'
END
ELSE
BEGIN
    PRINT 'El índice IX_ORDEFABR_IdSubOr ya existe.'
END
GO

-- Verificar estadísticas
PRINT 'Actualizando estadísticas de ORDEFABR...'
UPDATE STATISTICS [dbo].[ORDEFABR]
GO

-- Consulta de prueba para verificar performance
PRINT 'Ejecutando consulta de prueba...'
SET STATISTICS TIME ON
SET STATISTICS IO ON

SELECT TOP 10
    Id, FechGen, IdSubOr, CanProy, CanApro, Id_Orfa, StatuOrf,
    Cod_Exte, Unidad, FechEnSol, FechLan, FechCierr, Descrip, Cod_Inte
FROM ORDEFABR
WHERE IdSubOr LIKE 'DL-%'
ORDER BY IdSubOr DESC

SET STATISTICS TIME OFF
SET STATISTICS IO OFF
GO

PRINT 'Script completado.'
