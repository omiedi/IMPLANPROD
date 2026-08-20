SET NOCOUNT ON;

BEGIN TRY
    BEGIN TRAN;

    IF OBJECT_ID('dbo.SEG_MODULO', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.SEG_MODULO
        (
            Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
            Codigo NVARCHAR(80) NOT NULL,
            Nombre NVARCHAR(150) NOT NULL,
            Descripcion NVARCHAR(300) NULL,
            Orden INT NOT NULL CONSTRAINT DF_SEG_MODULO_Orden DEFAULT(0),
            Activo BIT NOT NULL CONSTRAINT DF_SEG_MODULO_Activo DEFAULT(1),
            Codiempr INT NOT NULL CONSTRAINT DF_SEG_MODULO_Codiempr DEFAULT(0),
            AddRecord DATETIME2 NULL CONSTRAINT DF_SEG_MODULO_AddRecord DEFAULT(SYSDATETIME()),
            LastUpdate DATETIME2 NULL CONSTRAINT DF_SEG_MODULO_LastUpdate DEFAULT(SYSDATETIME()),
            CodiEmprNet INT NULL CONSTRAINT DF_SEG_MODULO_CodiEmprNet DEFAULT(0),
            NumUsuar INT NULL CONSTRAINT DF_SEG_MODULO_NumUsuar DEFAULT(0)
        );

        CREATE UNIQUE INDEX UX_SEG_MODULO_Codigo ON dbo.SEG_MODULO(Codigo);
        CREATE INDEX IX_SEG_MODULO_Orden ON dbo.SEG_MODULO(Orden);
    END;

    IF OBJECT_ID('dbo.SEG_PAGINA', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.SEG_PAGINA
        (
            Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
            ModuloId INT NOT NULL,
            Nombre NVARCHAR(150) NOT NULL,
            Ruta NVARCHAR(200) NOT NULL,
            Componente NVARCHAR(160) NULL,
            EsConfiguracion BIT NOT NULL CONSTRAINT DF_SEG_PAGINA_EsConfiguracion DEFAULT(0),
            Orden INT NOT NULL CONSTRAINT DF_SEG_PAGINA_Orden DEFAULT(0),
            Activo BIT NOT NULL CONSTRAINT DF_SEG_PAGINA_Activo DEFAULT(1),
            Codiempr INT NOT NULL CONSTRAINT DF_SEG_PAGINA_Codiempr DEFAULT(0),
            AddRecord DATETIME2 NULL CONSTRAINT DF_SEG_PAGINA_AddRecord DEFAULT(SYSDATETIME()),
            LastUpdate DATETIME2 NULL CONSTRAINT DF_SEG_PAGINA_LastUpdate DEFAULT(SYSDATETIME()),
            CodiEmprNet INT NULL CONSTRAINT DF_SEG_PAGINA_CodiEmprNet DEFAULT(0),
            NumUsuar INT NULL CONSTRAINT DF_SEG_PAGINA_NumUsuar DEFAULT(0),
            CONSTRAINT FK_SEG_PAGINA_SEG_MODULO FOREIGN KEY (ModuloId) REFERENCES dbo.SEG_MODULO(Id)
        );

        CREATE UNIQUE INDEX UX_SEG_PAGINA_Ruta ON dbo.SEG_PAGINA(Ruta);
        CREATE INDEX IX_SEG_PAGINA_Modulo_Orden ON dbo.SEG_PAGINA(ModuloId, Orden);
    END;

    IF OBJECT_ID('dbo.SEG_USUARIO_PAGINA', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.SEG_USUARIO_PAGINA
        (
            Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
            UsuarioId INT NOT NULL,
            PaginaId INT NOT NULL,
            Permitido BIT NOT NULL CONSTRAINT DF_SEG_USUARIO_PAGINA_Permitido DEFAULT(1),
            Codiempr INT NOT NULL CONSTRAINT DF_SEG_USUARIO_PAGINA_Codiempr DEFAULT(0),
            AddRecord DATETIME2 NULL CONSTRAINT DF_SEG_USUARIO_PAGINA_AddRecord DEFAULT(SYSDATETIME()),
            LastUpdate DATETIME2 NULL CONSTRAINT DF_SEG_USUARIO_PAGINA_LastUpdate DEFAULT(SYSDATETIME()),
            CodiEmprNet INT NULL CONSTRAINT DF_SEG_USUARIO_PAGINA_CodiEmprNet DEFAULT(0),
            NumUsuar INT NULL CONSTRAINT DF_SEG_USUARIO_PAGINA_NumUsuar DEFAULT(0),
            CONSTRAINT FK_SEG_USUARIO_PAGINA_USUARIO FOREIGN KEY (UsuarioId) REFERENCES dbo.Usuarios(Id),
            CONSTRAINT FK_SEG_USUARIO_PAGINA_PAGINA FOREIGN KEY (PaginaId) REFERENCES dbo.SEG_PAGINA(Id)
        );

        CREATE UNIQUE INDEX UX_SEG_USUARIO_PAGINA_Usuario_Pagina ON dbo.SEG_USUARIO_PAGINA(UsuarioId, PaginaId);
    END;

    IF OBJECT_ID('dbo.SEG_AUDITORIA_PERFIL', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.SEG_AUDITORIA_PERFIL
        (
            Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
            UsuarioObjetivoId INT NOT NULL,
            PerfilAnteriorId INT NULL,
            PerfilNuevoId INT NOT NULL,
            ModificadoPorUsuarioId INT NOT NULL,
            EsCambioSuperAdmin BIT NOT NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_EsCambioSuperAdmin DEFAULT(0),
            Observacion NVARCHAR(400) NULL,
            FechaAccion DATETIME2 NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_FechaAccion DEFAULT(SYSDATETIME()),
            Codiempr INT NOT NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_Codiempr DEFAULT(0),
            AddRecord DATETIME2 NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_AddRecord DEFAULT(SYSDATETIME()),
            LastUpdate DATETIME2 NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_LastUpdate DEFAULT(SYSDATETIME()),
            CodiEmprNet INT NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_CodiEmprNet DEFAULT(0),
            NumUsuar INT NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_NumUsuar DEFAULT(0),
            CONSTRAINT FK_SEG_AUDITORIA_PERFIL_UsuarioObjetivo FOREIGN KEY (UsuarioObjetivoId) REFERENCES dbo.Usuarios(Id),
            CONSTRAINT FK_SEG_AUDITORIA_PERFIL_PerfilAnterior FOREIGN KEY (PerfilAnteriorId) REFERENCES dbo.Perfiles(Id),
            CONSTRAINT FK_SEG_AUDITORIA_PERFIL_PerfilNuevo FOREIGN KEY (PerfilNuevoId) REFERENCES dbo.Perfiles(Id),
            CONSTRAINT FK_SEG_AUDITORIA_PERFIL_ModificadoPor FOREIGN KEY (ModificadoPorUsuarioId) REFERENCES dbo.Usuarios(Id)
        );

        CREATE INDEX IX_SEG_AUDITORIA_PERFIL_Usuario_Fecha ON dbo.SEG_AUDITORIA_PERFIL(UsuarioObjetivoId, FechaAccion DESC);
    END;

    IF NOT EXISTS (SELECT 1 FROM dbo.Perfiles WHERE UPPER(LTRIM(RTRIM(Nombre))) = 'SUPERADMIN')
    BEGIN
        INSERT INTO dbo.Perfiles
        (
            Nombre,
            Descripcion,
            AddRecord,
            LastUpdate,
            CodiEmprNet,
            NumUsuar
        )
        VALUES
        (
            'SUPERADMIN',
            'Acceso total. Unico perfil que puede asignar/revocar SUPERADMIN.',
            SYSDATETIME(),
            SYSDATETIME(),
            0,
            0
        );
    END;

    IF NOT EXISTS (SELECT 1 FROM dbo.Perfiles WHERE UPPER(LTRIM(RTRIM(Nombre))) = 'ADMINISTRADOR')
    BEGIN
        INSERT INTO dbo.Perfiles (Nombre, Descripcion, AddRecord, LastUpdate, CodiEmprNet, NumUsuar)
        VALUES ('ADMINISTRADOR', 'Acceso a todo menos configuracion de seguridad.', SYSDATETIME(), SYSDATETIME(), 0, 0);
    END;

    IF NOT EXISTS (SELECT 1 FROM dbo.Perfiles WHERE UPPER(LTRIM(RTRIM(Nombre))) = 'USUARIO')
    BEGIN
        INSERT INTO dbo.Perfiles (Nombre, Descripcion, AddRecord, LastUpdate, CodiEmprNet, NumUsuar)
        VALUES ('USUARIO', 'Acceso a rutas explicitamente asignadas.', SYSDATETIME(), SYSDATETIME(), 0, 0);
    END;

    COMMIT;
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK;

    DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
    DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
    DECLARE @ErrorState INT = ERROR_STATE();

    RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;

-- Validación rápida post script
SELECT COUNT(1) AS TotalSuperAdminPerfil
FROM dbo.Perfiles
WHERE UPPER(LTRIM(RTRIM(Nombre))) = 'SUPERADMIN';
