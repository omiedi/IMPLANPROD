-- =============================================
-- Script: Agregar derecho "InCompra" para control de acceso al módulo de Compras
-- Fecha: 2026-04-13
-- Descripción: Agrega el derecho en INDI_APLI para controlar quién puede ingresar al módulo de Compras
-- =============================================

-- 1. Verificar si ya existe el derecho en INDI_APLI
IF NOT EXISTS (SELECT 1 FROM INDI_APLI WHERE COD_APLI = 'InCompra')
BEGIN
    -- Insertar el nuevo derecho en INDI_APLI
    INSERT INTO INDI_APLI (COD_APLI, DES_APLI)
    VALUES ('InCompra', 'Ingreso de Orden de Compras')
    
    PRINT 'Derecho "InCompra" agregado correctamente en INDI_APLI'
END
ELSE
BEGIN
    PRINT 'El derecho "InCompra" ya existe en INDI_APLI'
END
GO

-- 2. Asignar el derecho a usuarios ADMINISTRADORES automáticamente
-- Obtener todos los usuarios con perfil ADMINISTRADOR
DECLARE @PerfilAdminId INT

-- Buscar el ID del perfil ADMINISTRADOR
SELECT @PerfilAdminId = Id 
FROM Perfiles 
WHERE UPPER(Nombre) = 'ADMINISTRADOR'

IF @PerfilAdminId IS NOT NULL
BEGIN
    -- Insertar derechos para todos los usuarios administradores que no lo tengan
    INSERT INTO DER_ACCE (NUM_USER, COD_APLI, DER_READ, DER_WRITE, CodiEmprNet, NumUsuar, AddRecord, LastUpdate)
    SELECT 
        u.IdFlexoft,
        'InCompra',
        1, -- DER_READ = 1 (lectura)
        1, -- DER_WRITE = 1 (escritura)
        0, -- CodiEmprNet = 0
        u.IdFlexoft, -- NumUsuar
        GETDATE(), -- AddRecord
        GETDATE() -- LastUpdate
    FROM Usuarios u
    WHERE u.PerfilId = @PerfilAdminId
      AND NOT EXISTS (
          SELECT 1 
          FROM DER_ACCE d 
          WHERE d.NUM_USER = u.IdFlexoft 
            AND d.COD_APLI = 'InCompra'
      )
    
    PRINT 'Derechos "InCompra" asignados a usuarios ADMINISTRADORES'
END
ELSE
BEGIN
    PRINT 'No se encontró el perfil ADMINISTRADOR'
END
GO

-- 3. Verificar los registros creados
SELECT 
    'INDI_APLI' AS Tabla,
    COD_APLI,
    DES_APLI
FROM INDI_APLI
WHERE COD_APLI = 'InCompra'

UNION ALL

SELECT 
    'DER_ACCE' AS Tabla,
    CAST(NUM_USER AS VARCHAR(32)) AS COD_APLI,
    'User: ' + CAST(NUM_USER AS VARCHAR) + ' - Read: ' + CAST(DER_READ AS VARCHAR) + ' - Write: ' + CAST(DER_WRITE AS VARCHAR) AS DES_APLI
FROM DER_ACCE
WHERE COD_APLI = 'InCompra'
ORDER BY Tabla, COD_APLI
GO
