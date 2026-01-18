-- =============================================
-- Script: Diagnóstico de O.F. Limpieza
-- Fecha: 2024-12-16
-- Descripción: Verifica cantidad de registros DL- y performance de consultas
-- =============================================

USE [IMPLANPROD]
GO

PRINT '========================================='
PRINT 'DIAGNÓSTICO DE O.F. LIMPIEZA (DL-)'
PRINT '========================================='
PRINT ''

-- 1. Contar registros DL-
PRINT '1. CANTIDAD DE REGISTROS DL-:'
SELECT 
    COUNT(*) AS [Total_Registros_DL],
    COUNT(CASE WHEN StatuOrf < 3 THEN 1 END) AS [Abiertas],
    COUNT(CASE WHEN StatuOrf >= 3 THEN 1 END) AS [Cerradas]
FROM ORDEFABR
WHERE IdSubOr LIKE 'DL-%'
GO

-- 2. Verificar índices existentes
PRINT ''
PRINT '2. ÍNDICES EN ORDEFABR:'
SELECT 
    i.name AS [Nombre_Indice],
    i.type_desc AS [Tipo],
    COL_NAME(ic.object_id, ic.column_id) AS [Columna],
    ic.is_included_column AS [Es_Include]
FROM sys.indexes i
INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
WHERE i.object_id = OBJECT_ID('ORDEFABR')
    AND (i.name LIKE '%IdSubOr%' OR i.is_primary_key = 1)
ORDER BY i.name, ic.key_ordinal
GO

-- 3. Verificar fragmentación de índices
PRINT ''
PRINT '3. FRAGMENTACIÓN DE ÍNDICES:'
SELECT 
    i.name AS [Nombre_Indice],
    ps.avg_fragmentation_in_percent AS [Fragmentacion_%],
    ps.page_count AS [Paginas]
FROM sys.dm_db_index_physical_stats(DB_ID(), OBJECT_ID('ORDEFABR'), NULL, NULL, 'LIMITED') ps
INNER JOIN sys.indexes i ON ps.object_id = i.object_id AND ps.index_id = i.index_id
WHERE ps.avg_fragmentation_in_percent > 10
ORDER BY ps.avg_fragmentation_in_percent DESC
GO

-- 4. Probar performance de consulta actual
PRINT ''
PRINT '4. PERFORMANCE DE CONSULTA (TOP 10 DL-):'
SET STATISTICS TIME ON
SET STATISTICS IO ON

SELECT TOP 10
    Id, FechGen, IdSubOr, CanProy, CanApro, Id_Orfa, StatuOrf,
    Cod_Exte, Unidad, FechEnSol, FechLan, FechCierr, Descrip, Cod_Inte
FROM ORDEFABR WITH (NOLOCK)
WHERE IdSubOr LIKE 'DL-%'
ORDER BY IdSubOr DESC

SET STATISTICS TIME OFF
SET STATISTICS IO OFF
GO

-- 5. Ver plan de ejecución estimado
PRINT ''
PRINT '5. PLAN DE EJECUCIÓN:'
SET SHOWPLAN_TEXT ON
GO

SELECT TOP 10
    Id, FechGen, IdSubOr, CanProy, CanApro, Id_Orfa, StatuOrf,
    Cod_Exte, Unidad, FechEnSol, FechLan, FechCierr, Descrip, Cod_Inte
FROM ORDEFABR
WHERE IdSubOr LIKE 'DL-%'
ORDER BY IdSubOr DESC
GO

SET SHOWPLAN_TEXT OFF
GO

-- 6. Verificar bloqueos activos
PRINT ''
PRINT '6. BLOQUEOS ACTIVOS:'
SELECT 
    request_session_id AS [Session_ID],
    resource_type AS [Tipo_Recurso],
    resource_database_id AS [DB_ID],
    resource_description AS [Descripcion],
    request_mode AS [Modo],
    request_status AS [Estado]
FROM sys.dm_tran_locks
WHERE resource_database_id = DB_ID()
    AND resource_type IN ('OBJECT', 'PAGE', 'KEY')
GO

PRINT ''
PRINT '========================================='
PRINT 'DIAGNÓSTICO COMPLETADO'
PRINT '========================================='
