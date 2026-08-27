-- ===================================================================
-- REGISTRO RBAC PARA EL MODULO SEGUIMIENTO DE SOLICITUDES
-- IMPORTANTE: Ajustar @ModuloId segun el modulo correspondiente o
-- el modulo de Configuracion para la ruta admin.
-- ===================================================================

-- ===================================================================
-- 1. RUTA PUBLICA/VISTA PROPIA (visible y accesible por todos los usuarios)
-- ===================================================================

DECLARE @ModuloUsuarios INT = (SELECT TOP 1 Id FROM dbo.SEG_MODULO WHERE Codigo = 'USUARIOS' OR Nombre = 'Usuarios' ORDER BY Id);
DECLARE @ModuloConfig INT = (SELECT TOP 1 Id FROM dbo.SEG_MODULO WHERE Codigo = 'CONFIGURACION' OR Nombre = 'Configuración' ORDER BY Id);

-- Ajuste manual si no existe modulo correspondiente:
-- SET @ModuloUsuarios = 1; -- reemplazar por el ID real
-- SET @ModuloConfig = 2;   -- reemplazar por el ID real

IF NOT EXISTS (SELECT 1 FROM dbo.SEG_PAGINA WHERE Ruta = '/seguimiento-solicitudes')
BEGIN
    INSERT INTO dbo.SEG_PAGINA (ModuloId, Nombre, Ruta, Componente, EsConfiguracion, Orden, Activo, Codiempr, AddRecord, LastUpdate, CodiEmprNet, NumUsuar)
    VALUES (@ModuloUsuarios, 'Seguimiento de Solicitudes', '/seguimiento-solicitudes', 'Client/Pages/Solicitudes/SeguimientoSolicitudes.razor', 0, 0, 1, 0, SYSDATETIME(), SYSDATETIME(), 0, 0);
END;

IF NOT EXISTS (SELECT 1 FROM dbo.SEG_PAGINA WHERE Ruta = '/admin/seguimiento-solicitudes')
BEGIN
    INSERT INTO dbo.SEG_PAGINA (ModuloId, Nombre, Ruta, Componente, EsConfiguracion, Orden, Activo, Codiempr, AddRecord, LastUpdate, CodiEmprNet, NumUsuar)
    VALUES (@ModuloConfig, 'Seguimiento de Solicitudes (Admin)', '/admin/seguimiento-solicitudes', 'Client/Pages/Solicitudes/SeguimientoSolicitudes.razor', 1, 0, 1, 0, SYSDATETIME(), SYSDATETIME(), 0, 0);
END;

-- ===================================================================
-- 2. ASIGNACION A TODOS LOS USUARIOS ACTIVOS PARA LA RUTA PROPIA
--    La ruta admin debe asignarse manualmente por admin/superadmin.
-- ===================================================================

DECLARE @PaginaUsuario INT = (SELECT Id FROM dbo.SEG_PAGINA WHERE Ruta = '/seguimiento-solicitudes');

IF @PaginaUsuario IS NOT NULL
BEGIN
    INSERT INTO dbo.SEG_USUARIO_PAGINA (UsuarioId, PaginaId, Permitido, AddRecord, LastUpdate, CodiEmprNet, NumUsuar)
    SELECT u.Id, @PaginaUsuario, 1, SYSDATETIME(), SYSDATETIME(), 0, 0
    FROM dbo.Usuarios u
    WHERE u.Activo = 1
      AND NOT EXISTS (
          SELECT 1 FROM dbo.SEG_USUARIO_PAGINA sup
          WHERE sup.UsuarioId = u.Id AND sup.PaginaId = @PaginaUsuario
      );
END;

-- ===================================================================
-- 3. ASIGNACION A SUPERADMIN/ADMIN PARA LA RUTA ADMINISTRATIVA
-- ===================================================================

DECLARE @PaginaAdmin INT = (SELECT Id FROM dbo.SEG_PAGINA WHERE Ruta = '/admin/seguimiento-solicitudes');

IF @PaginaAdmin IS NOT NULL
BEGIN
    INSERT INTO dbo.SEG_USUARIO_PAGINA (UsuarioId, PaginaId, Permitido, AddRecord, LastUpdate, CodiEmprNet, NumUsuar)
    SELECT u.Id, @PaginaAdmin, 1, SYSDATETIME(), SYSDATETIME(), 0, 0
    FROM dbo.Usuarios u
    INNER JOIN dbo.Perfiles p ON u.PerfilId = p.Id
    WHERE u.Activo = 1
      AND (UPPER(p.Nombre) IN ('SUPERADMIN','ADMINISTRADOR','ADMIN'))
      AND NOT EXISTS (
          SELECT 1 FROM dbo.SEG_USUARIO_PAGINA sup
          WHERE sup.UsuarioId = u.Id AND sup.PaginaId = @PaginaAdmin
      );
END;
