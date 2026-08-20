-- =============================================
-- Script: AddDerechoVerCostosOC.sql
-- Descripción: Agrega el derecho "VerCostosOC" para controlar quién puede ver y editar costos de órdenes de compra
-- Fecha: 2026-04-13
-- =============================================

-- 1. Verificar si ya existe el derecho en INDI_APLI
IF NOT EXISTS (SELECT 1 FROM INDI_APLI WHERE COD_APLI = 'VerCostosOC')
BEGIN
    -- Insertar el nuevo derecho en INDI_APLI
    INSERT INTO INDI_APLI (COD_APLI, DES_APLI)
    VALUES ('VerCostosOC', 'Ver y Editar Costos de Órdenes de Compra')
    
    PRINT 'Derecho "VerCostosOC" agregado correctamente en INDI_APLI'
END
ELSE
BEGIN
    PRINT 'El derecho "VerCostosOC" ya existe en INDI_APLI'
END
GO

-- 2. Asignar el derecho a usuarios ADMINISTRADORES automáticamente
-- Obtener todos los usuarios con perfil ADMINISTRADOR
DECLARE @PerfilAdminId INT
SELECT @PerfilAdminId = IDPERFIL FROM PERFILES WHERE UPPER(NOMPERFIL) LIKE '%ADMIN%'

IF @PerfilAdminId IS NOT NULL
BEGIN
    -- Insertar permisos para usuarios ADMINISTRADORES que no los tengan
    INSERT INTO DER_ACCE (NUM_USER, COD_APLI, DER_READ, DER_WRITE)
    SELECT 
        u.idflexoft,
        'VerCostosOC',
        2, -- Lectura completa
        2  -- Escritura completa
    FROM usuarios u
    WHERE u.IDPERFIL = @PerfilAdminId
    AND NOT EXISTS (
        SELECT 1 
        FROM DER_ACCE da 
        WHERE da.NUM_USER = u.idflexoft 
        AND da.COD_APLI = 'VerCostosOC'
    )
    
    PRINT 'Derecho "VerCostosOC" asignado a usuarios ADMINISTRADORES'
END
ELSE
BEGIN
    PRINT 'No se encontró el perfil ADMINISTRADOR'
END
GO

-- 3. Verificar la creación
SELECT 
    'INDI_APLI' AS Tabla,
    COD_APLI,
    DES_APLI
FROM INDI_APLI
WHERE COD_APLI = 'VerCostosOC'

UNION ALL

SELECT 
    'DER_ACCE' AS Tabla,
    CAST(NUM_USER AS VARCHAR(32)) AS COD_APLI,
    'User: ' + CAST(NUM_USER AS VARCHAR) + ' - Read: ' + CAST(DER_READ AS VARCHAR) + ' - Write: ' + CAST(DER_WRITE AS VARCHAR) AS DES_APLI
FROM DER_ACCE
WHERE COD_APLI = 'VerCostosOC'
ORDER BY Tabla, COD_APLI
GO

PRINT '============================================='
PRINT 'Script completado exitosamente'
PRINT 'Derecho: VerCostosOC'
PRINT 'Descripción: Ver y Editar Costos de Órdenes de Compra'
PRINT 'Usuarios con acceso: ADMINISTRADORES (nivel 2/2)'
PRINT '============================================='
GO
