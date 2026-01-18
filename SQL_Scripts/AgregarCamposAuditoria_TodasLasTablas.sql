-- =============================================
-- Script para agregar campos de auditoría y timestamp
-- a todas las tablas especificadas
-- Fecha: 2024
-- =============================================
-- IMPORTANTE: Ejecutar este script en SQL Server Management Studio
-- Los campos se agregan con valores por defecto según especificación:
-- AddRecord: DATETIME NOT NULL DEFAULT (GETDATE())
-- LastUpdate: DATETIME NOT NULL DEFAULT (GETDATE())
-- CodiEmprNet: INT NOT NULL DEFAULT (101)
-- NumUsuar: INT NOT NULL DEFAULT (10)
-- =============================================

USE [TU_BASE_DE_DATOS]
GO

-- =============================================
-- 1. CATOPER
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Catoper]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Catoper]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Catoper';
END
GO

-- =============================================
-- 2. CITIES
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Cities]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Cities]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Cities';
END
GO

-- =============================================
-- 3. CONDPAG
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[CONDPAG]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[CONDPAG]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a CONDPAG';
END
GO

-- =============================================
-- 4. COUNTRIES
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Countries]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Countries]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Countries';
END
GO

-- =============================================
-- 5. DEPENDENCIAS
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Dependencias]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Dependencias]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Dependencias';
END
GO

-- =============================================
-- 6. DESTINORECHAZORECEPCION
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[DestinoRechazoRecepcion]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[DestinoRechazoRecepcion]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a DestinoRechazoRecepcion';
END
GO

-- =============================================
-- 7. DEUDOR12
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Deudor12]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Deudor12]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Deudor12';
END
GO

-- =============================================
-- 8. EMAILCONFIGURATION
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[EmailConfiguration]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[EmailConfiguration]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a EmailConfiguration';
END
GO

-- =============================================
-- 9. EMPRESAS
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Empresas]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Empresas]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Empresas';
END
GO

-- =============================================
-- 10. FAMILIA
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Familia]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Familia]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Familia';
END
GO

-- =============================================
-- 11. FLEXCRL
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[FLEXCRL]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[FLEXCRL]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a FLEXCRL';
END
GO

-- =============================================
-- 12. FORMULAS
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Formulas]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Formulas]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Formulas';
END
GO

-- =============================================
-- 13. GRUPIVA
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Grupiva]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Grupiva]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Grupiva';
END
GO

-- =============================================
-- 14. MASTER
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Master]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Master]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Master';
END
GO

-- =============================================
-- 15. MOTIVOSRECHAZORECEPCION
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[MotivosRechazoRecepcion]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[MotivosRechazoRecepcion]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a MotivosRechazoRecepcion';
END
GO

-- =============================================
-- 16. MOVITRANSTO
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[MOVITRANSTO]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[MOVITRANSTO]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a MOVITRANSTO';
END
GO

-- =============================================
-- 17. OPERFAB
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[OPERFAB]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[OPERFAB]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a OPERFAB';
END
GO

-- =============================================
-- 18. ORDEFABR
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Ordefabr]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Ordefabr]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Ordefabr';
END
GO

-- =============================================
-- 19. ORFAPEN
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ORFAPEN]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[ORFAPEN]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101);
    -- NOTA: NumUsuar ya existe en ORFAPEN, no se agrega
    PRINT 'Campos de auditoría agregados a ORFAPEN (NumUsuar ya existía)';
END
GO

-- =============================================
-- 20. PADMAQ
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[PADMAQ]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[PADMAQ]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a PADMAQ';
END
GO

-- =============================================
-- 21. PEDIDOS_AUDITORIA
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[PEDIDOS_AUDITORIA]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[PEDIDOS_AUDITORIA]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a PEDIDOS_AUDITORIA';
END
GO

-- =============================================
-- 22. PERFILES
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Perfiles]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Perfiles]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Perfiles';
END
GO

-- =============================================
-- 23. PERMISOS
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Permisos]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Permisos]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Permisos';
END
GO

-- =============================================
-- 24. PROVED12
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Proved12]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Proved12]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Proved12';
END
GO

-- =============================================
-- 25. RESERVA
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Reserva]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Reserva]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Reserva';
END
GO

-- =============================================
-- 26. SECOPER
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[SECOPER]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[SECOPER]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a SECOPER';
END
GO

-- =============================================
-- 27. SECTORES
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Sectores]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Sectores]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Sectores';
END
GO

-- =============================================
-- 28. STATES
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[States]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[States]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a States';
END
GO

-- =============================================
-- 29. TADEPOSI
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Tadeposi]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Tadeposi]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Tadeposi';
END
GO

-- =============================================
-- 30. TAMONED
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Tamoned]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Tamoned]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Tamoned';
END
GO

-- =============================================
-- 31. TATIMAT
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Tatimat]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[Tatimat]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Tatimat';
END
GO

-- =============================================
-- 32. TRANSPORTES
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[TRANSPORTES]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[TRANSPORTES]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a TRANSPORTES';
END
GO

-- =============================================
-- 33. UNIDADMEDIDAS
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[UnidadMedidas]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[UnidadMedidas]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a UnidadMedidas';
END
GO

-- =============================================
-- 34. UPLOADMASTER
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[UploadMaster]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[UploadMaster]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a UploadMaster';
END
GO

-- =============================================
-- 35. USUARIOS
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Usuarios]') AND name = 'CodiEmprNet')
BEGIN
    ALTER TABLE [dbo].[Usuarios]
    ADD [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a Usuarios (AddRecord y LastUpdate ya existían)';
END
GO

-- =============================================
-- 36. USUARIOSSECTORPROD
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[UsuariosSectorProd]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[UsuariosSectorProd]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a UsuariosSectorProd';
END
GO

-- =============================================
-- 37. VENDEDOR
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[VENDEDOR]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[VENDEDOR]
    ADD [AddRecord] DATETIME NOT NULL DEFAULT (GETDATE()),
        [LastUpdate] DATETIME NOT NULL DEFAULT (GETDATE()),
        [CodiEmprNet] INT NOT NULL DEFAULT (101),
        [NumUsuar] INT NOT NULL DEFAULT (10);
    PRINT 'Campos de auditoría agregados a VENDEDOR';
END
GO

PRINT '=============================================';
PRINT 'Script completado exitosamente';
PRINT 'Total de tablas procesadas: 37';
PRINT '=============================================';
