using DocumentFormat.OpenXml.Office2010.Excel;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.EntityFrameworkCore;
using System.Security.Cryptography;
using System.Text;

namespace IMPLANPROD.Server.Data
{
    public class SeedDb
    {
        private readonly DataContext _context;
        private readonly ILogger<SeedDb> _logger;
        private readonly GeorefService _georefService;
        private readonly ILoggerFactory _loggerFactory;

        public SeedDb(DataContext context, ILogger<SeedDb> logger, GeorefService georefService, ILoggerFactory loggerFactory)
        {
            _context = context;
            _logger = logger;
            _georefService = georefService;
            _loggerFactory = loggerFactory;
        }

        public async Task SeedAsync()
        {
            await _context.Database.EnsureCreatedAsync();

            // Verificar si la tabla EmailConfiguration existe
            try
            {
                _ = await _context.Set<EmailConfiguration>().AnyAsync();
            }
            catch
            {
                // Si hay una excepción, significa que la tabla no existe
                await _context.Database.ExecuteSqlRawAsync(@"
                    CREATE TABLE EmailConfiguration (
                        Id INT IDENTITY(1,1) PRIMARY KEY,
                        EmailAddress NVARCHAR(450) NOT NULL,
                        EncryptedPassword NVARCHAR(MAX) NOT NULL,
                        DisplayName NVARCHAR(256) NOT NULL,
                        LastModified DATETIME2 NOT NULL
                    );

                    CREATE UNIQUE INDEX IX_EmailConfiguration_EmailAddress ON EmailConfiguration(EmailAddress);
                ");
            }

            // Verificar si la tabla Usuarios existe
            try
            {
                _ = await _context.Usuarios.AnyAsync();
            }
            catch
            {
                // Si hay una excepción, significa que la tabla no existe
                await _context.Database.ExecuteSqlRawAsync(@"
                    CREATE TABLE Usuarios (
                        Id INT IDENTITY(1,1) PRIMARY KEY,
                        NombreUsuario NVARCHAR(450) NOT NULL,
                        Email NVARCHAR(450) NOT NULL,
                        Password NVARCHAR(MAX) NOT NULL,
                        FechaRegistro DATETIME2 NOT NULL,
                        Activo BIT NOT NULL,
                        FotoUrl NVARCHAR(MAX) NULL,
                        PerfilId INT NOT NULL,
                        EmpresaId INT NOT NULL,
                        DependenciaId INT NOT NULL,
                        SectorId INT NOT NULL,
                        CountryId INT NOT NULL,
                        StateId INT NOT NULL,
                        CityId INT NOT NULL,
                        IntentosLogin INT NOT NULL,
                        Bloqueado BIT NOT NULL,
                        UltimoAcceso DATETIME2 NOT NULL,
                        RefreshToken NVARCHAR(MAX) NULL,
                        RefreshTokenExpiryTime DATETIME2 NULL,
                        -- Se especifica ON DELETE NO ACTION en las claves foráneas para evitar ciclos o rutas múltiples en cascada
                        CONSTRAINT FK_Usuarios_Perfiles FOREIGN KEY (PerfilId) REFERENCES Perfiles(Id),
                        CONSTRAINT FK_Usuarios_Empresas FOREIGN KEY (EmpresaId) REFERENCES Empresas(Id),
                        CONSTRAINT FK_Usuarios_Dependencias FOREIGN KEY (DependenciaId) REFERENCES Dependencias(Id),
                        CONSTRAINT FK_Usuarios_Sectores FOREIGN KEY (SectorId) REFERENCES Sectores(Id),
                        CONSTRAINT FK_Usuarios_Countries FOREIGN KEY (CountryId) REFERENCES Countries(Id) ON DELETE NO ACTION,
                        CONSTRAINT FK_Usuarios_States FOREIGN KEY (StateId) REFERENCES States(Id) ON DELETE NO ACTION,
                        CONSTRAINT FK_Usuarios_Cities FOREIGN KEY (CityId) REFERENCES Cities(Id) ON DELETE NO ACTION
                    );
                    
                    CREATE UNIQUE INDEX IX_Usuarios_NombreUsuario ON Usuarios(NombreUsuario);
                    CREATE UNIQUE INDEX IX_Usuarios_Email ON Usuarios(Email);
                ");
            }

            // Verificar si la tabla Deudor12 existe usando SQL directo
            bool deudor12Exists = false;
            try
            {
                // Consulta para verificar si la tabla existe en la base de datos
                var result = await _context.Database.ExecuteSqlRawAsync(
                    "IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Deudor12') SELECT 1 ELSE SELECT 0");
                deudor12Exists = result > 0;
            }
            catch
            {
                // Si hay un error, asumimos que la tabla no existe
                deudor12Exists = false;
            }

            // Solo creamos la tabla si no existe
            if (!deudor12Exists)
            {
                try
                {
                    await _context.Database.ExecuteSqlRawAsync(@"
                        CREATE TABLE [dbo].[Deudor12](
                            [Id] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
                            [Nume_Cli] [int] NULL,
                            [Lentre0_Cli] [nvarchar](255) NULL,
                            [Lentre1_Cli] [nvarchar](255) NULL,
                            [NTrans_Cli] [decimal](18,2) NULL,
                            [Lentre2_Cli] [nvarchar](255) NULL,
                            [Hora_Ent] [nvarchar](255) NULL,
                            [Zona_Cli] [nvarchar](255) NULL,
                            [Zona_Rep] [nvarchar](255) NULL,
                            [Gru_Div] [decimal](18,2) NULL,
                            [Codi_Cli] [nvarchar](255) NULL,
                            [Raso_Cli] [nvarchar](255) NULL,
                            [Fant_Cli] [nvarchar](255) NULL,
                            [Stat_Cli] [smallint] NULL,
                            [Domi_Cli] [nvarchar](255) NULL,
                            [Cpos_Cli] [nvarchar](255) NULL,
                            [Loca_Cli] [nvarchar](255) NULL,
                            [Pais_Cli] [nvarchar](255) NULL,
                            [Prov_Cli] [nvarchar](255) NULL,
                            [Tele_Cli] [nvarchar](255) NULL,
                            [Faxi_Cli] [nvarchar](255) NULL,
                            [Mail_Cli] [nvarchar](255) NULL,
                            [PagWeb] [nvarchar](255) NULL,
                            [Piva_Cli] [smallint] NULL,
                            [Cuit_Cli] [nvarchar](255) NULL,
                            [Nibr_Cli] [nvarchar](255) NULL,
                            [Prvd_Cli] [nvarchar](255) NULL,
                            [Ctac_Cli] [nvarchar](255) NULL,
                            [Hora_Cli] [nvarchar](255) NULL,
                            [Cpag_Cli] [decimal](18,2) NULL,
                            [Lpre_Cli] [decimal](18,2) NULL,
                            [Vend_Cli] [decimal](18,2) NULL,
                            [Tcpr_Cli] [smallint] NULL,
                            [CaIb_Cli] [nvarchar](255) NULL,
                            [Dsta_Cli] [nvarchar](255) NULL,
                            [Lcre_Cli] [decimal](18,2) NULL,
                            [Sald_Cli] [decimal](18,2) NULL,
                            [CuMa_Cli] [decimal](18,2) NULL,
                            [GrCo_Cli] [decimal](18,2) NULL,
                            [Nota_Cli] [ntext] NULL,
                            [Obser_Entre] [ntext] NULL,
                            [Recar_Embal] [decimal](18,2) NULL,
                            [Mar_Ya_Fac] [decimal](18,2) NULL,
                            [DePosIva] [nvarchar](255) NULL,
                            [CodiEmpr] [smallint] NULL,
                            [LCred_CCte] [decimal](18,2) NULL,
                            [EnvMaiDeuda] [decimal](18,2) NULL,
                            [Ctact_Cob] [nvarchar](48) NULL,
                            [Tele_Cob] [nvarchar](48) NULL,
                            [Mail_Cob] [nvarchar](128) NULL,
                            [DiasRec_Cob] [nvarchar](64) NULL,
                            [DiasPago_Cob] [nvarchar](64) NULL,
                            [Segui_Cob] [ntext] NULL,
                            [MailFactura_Cob] [nvarchar](128) NULL,
                            [CELU_CLI] [nvarchar](128) NULL,
                            [Lentre3_Cli] [nvarchar](10) NULL,
                            [JurisdiccioIIBB_CLI] [nvarchar](2) NULL,
                            [RegimenFacturaCredito] [smallint] NULL,
                            [FechAlta] [smalldatetime] NULL,
                            [ProvEntre_Cli] [nvarchar](2) NULL
                        ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
                    ");
                }
                catch (Exception ex)
                {
                    // Registrar el error pero continuar con el resto del proceso
                    Console.WriteLine($"Error al crear la tabla Deudor12: {ex.Message}");
                }
            }

            //// Ejecutar el seeder de datos geográficos de Argentina
            //var geoSeeder = new GeoArgentinaSeeder(_context, _georefService, _loggerFactory.CreateLogger<GeoArgentinaSeeder>());
            //await geoSeeder.SeedGeoDataAsync();

            // Resto de la inicialización
            await CheckEmpresasAsync();
            await CheckDependenciasAsync();
            await CheckSectoresAsync();
            await CheckPerfilesYPermisosAsync();
            await CheckUsuariosAsync();
            //await CheckUnidadMedidaAsync();
            await CheckEmailConfigurationAsync();
            await CheckProved12TableAsync();
        }

        // ------------------------------------------------------------------------
        // Método para crear la tabla Proved12 si no existe (solo SQL Server)
        // ------------------------------------------------------------------------
        private async Task CheckProved12TableAsync()
        {
            /*
               Este método verifica si la tabla Proved12 existe y la crea si no existe,
               usando la definición actual de la entidad Proved12.
               IMPORTANTE: Si cambias la entidad, actualiza también aquí la definición.
            */
            var sql = @"
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Proved12' AND xtype='U')
            BEGIN
                CREATE TABLE [dbo].[Proved12](
                    [Id] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
                    [Nume_Cli] [int] NULL,
                    Lcre_Cli DECIMAL(18,2) NULL,
                    CodiEmpr SMALLINT NULL,
                    Sald_Cli DECIMAL(18,2) NULL,
                    CuMa_Cli SMALLINT NULL,
                    GrCo_Cli SMALLINT NULL,
                    Nota_Cli NVARCHAR(MAX) NULL,
                    C_Ret NVARCHAR(10) NULL,
                    Posivat_Cli NVARCHAR(10) NULL,
                    Codi_Cli NVARCHAR(12) NULL,
                    Raso_Cli NVARCHAR(128) NOT NULL,
                    Fant_Cli NVARCHAR(32) NULL,
                    Stat_Cli SMALLINT NULL,
                    Domi_Cli NVARCHAR(64) NULL,
                    Cpos_Cli NVARCHAR(16) NULL,
                    Loca_Cli NVARCHAR(48) NULL,
                    Pais_Cli NVARCHAR(48) NULL,
                    Prov_Cli NVARCHAR(48) NULL,
                    Tele_Cli NVARCHAR(32) NULL,
                    Tele2_Cli NVARCHAR(32) NULL,
                    Faxi_Cli NVARCHAR(32) NULL,
                    Mail_Cli NVARCHAR(64) NULL,
                    Piva_Cli SMALLINT NULL,
                    Cuit_Cli NVARCHAR(24) NULL,
                    Nibr_Cli NVARCHAR(24) NULL,
                    Prvd_Cli NVARCHAR(24) NULL,
                    Ctac_Cli NVARCHAR(48) NULL,
                    Hora_Cli NVARCHAR(32) NULL,
                    Ord_Cheq NVARCHAR(48) NULL,
                    Cpag_Cli SMALLINT NULL,
                    CCom_Cli NVARCHAR(MAX) NULL,
                    Lpre_Cli SMALLINT NULL,
                    Vend_Cli SMALLINT NULL,
                    Tcpr_Cli SMALLINT NULL,
                    CaIb_Cli NVARCHAR(2) NULL,
                    Dsta_Cli NVARCHAR(32) NULL,
                    Calinteg DECIMAL(18,2) NULL,
                    ObsCompras NVARCHAR(MAX) NULL,
                    LuEntCompras NVARCHAR(MAX) NULL,
                    CategoriaIBCABA_Cli SMALLINT NULL,
                    CategoriaIBMISIONES_Cli SMALLINT NULL,
                    CategoriaIBTUCU_Cli SMALLINT NULL
                );
            END
            ";
            await _context.Database.ExecuteSqlRawAsync(sql);
        }

        //private async Task CheckCountriesAsync()
        //{
        //    if (!_context.Countries.Any())
        //    {
        //        _context.Countries.Add(new Country
        //        {
        //            Name = "Argentina",
        //            status = 0
        //        });

        //        await _context.SaveChangesAsync();
        //    }
        //}

        //private async Task CheckStatesAsync()
        //{
        //    if (!_context.States.Any())
        //    {
        //        var country = await _context.Countries.FirstOrDefaultAsync(c => c.Name == "Argentina");
        //        if (country != null)
        //        {
        //            var states = new List<State>
        //            {
        //                new State { CountryId = country.Id, Name = "Buenos Aires", status = 0 },
        //                new State { CountryId = country.Id, Name = "Córdoba", status = 0 },
        //                new State { CountryId = country.Id, Name = "Santa Fe", status = 0 }
        //            };

        //            await _context.States.AddRangeAsync(states);
        //            await _context.SaveChangesAsync();
        //        }
        //    }
        //}

        //private async Task CheckCitiesAsync()
        //{
        //    if (!_context.Cities.Any())
        //    {
        //        var states = await _context.States.ToListAsync();
        //        foreach (var state in states)
        //        {
        //            var cities = new List<City>();
        //            switch (state.Name)
        //            {
        //                case "Buenos Aires":
        //                    cities.Add(new City { StateId = state.Id, Name = "La Plata", status = 0 });
        //                    cities.Add(new City { StateId = state.Id, Name = "Mar del Plata", status = 0 });
        //                    break;
        //                case "Córdoba":
        //                    cities.Add(new City { StateId = state.Id, Name = "Córdoba Capital", status = 0 });
        //                    cities.Add(new City { StateId = state.Id, Name = "Villa María", status = 0 });
        //                    break;
        //                case "Santa Fe":
        //                    cities.Add(new City { StateId = state.Id, Name = "Rosario", status = 0 });
        //                    cities.Add(new City { StateId = state.Id, Name = "Santa Fe Capital", status = 0 });
        //                    break;
        //            }
        //            await _context.Cities.AddRangeAsync(cities);
        //        }
        //        await _context.SaveChangesAsync();
        //    }
        //}

        private async Task CheckEmpresasAsync()
        {
            if (!_context.Empresas.Any())
            {
                var empresas = new List<Empresa>
                {
                    new Empresa
                    {
                        razon = "IMPLAN San Luis",
                        fantasia = "IMPLAN",
                        CUIT = "30-99999999-9",
                        calle = "Belgrano",
                        numero = "1200",
                        localidad = "San Luis",
                        CP = "5700"
                    }
                };

                await _context.Empresas.AddRangeAsync(empresas);
                await _context.SaveChangesAsync();
            }
        }

        private async Task CheckDependenciasAsync()
        {
            if (!_context.Dependencias.Any())
            {
                var empresa = await _context.Empresas.FirstOrDefaultAsync();
                if (empresa != null)
                {
                    var dependencias = new List<Dependencia>
                    {
                        new Dependencia
                        {
                            EmpresaId = empresa.Id,
                            dependencia = "Administración"
                        },
                        new Dependencia
                        {
                            EmpresaId = empresa.Id,
                            dependencia = "Desarrollo"
                        },
                        new Dependencia
                        {
                            EmpresaId = empresa.Id,
                            dependencia = "Planificación"
                        }
                    };

                    await _context.Dependencias.AddRangeAsync(dependencias);
                    await _context.SaveChangesAsync();
                }
            }
        }

        private async Task CheckSectoresAsync()
        {
            if (!_context.Sectores.Any())
            {
                var dependencias = await _context.Dependencias.ToListAsync();
                foreach (var dependencia in dependencias)
                {
                    var sectores = new List<Sector>();
                    switch (dependencia.dependencia)
                    {
                        case "Administración":
                            sectores.Add(new Sector { DependenciaId = dependencia.Id, sector = "Recursos Humanos" });
                            sectores.Add(new Sector { DependenciaId = dependencia.Id, sector = "Contabilidad" });
                            break;
                        case "Desarrollo":
                            sectores.Add(new Sector { DependenciaId = dependencia.Id, sector = "Frontend" });
                            sectores.Add(new Sector { DependenciaId = dependencia.Id, sector = "Backend" });
                            break;
                        case "Planificación":
                            sectores.Add(new Sector { DependenciaId = dependencia.Id, sector = "Proyectos" });
                            sectores.Add(new Sector { DependenciaId = dependencia.Id, sector = "Estrategia" });
                            break;
                    }
                    await _context.Sectores.AddRangeAsync(sectores);
                }
                await _context.SaveChangesAsync();
            }
        }

        private async Task CheckPerfilesYPermisosAsync()
        {
            if (!_context.Perfiles.Any())
            {
                // Crear permisos básicos
                var permisos = new List<Permiso>
                {
                    new Permiso { Nombre = "Usuarios_Ver", Descripcion = "Ver usuarios", Clave = "USR_VIEW" },
                    new Permiso { Nombre = "Usuarios_Crear", Descripcion = "Crear usuarios", Clave = "USR_CREATE" },
                    new Permiso { Nombre = "Usuarios_Editar", Descripcion = "Editar usuarios", Clave = "USR_EDIT" },
                    new Permiso { Nombre = "Usuarios_Eliminar", Descripcion = "Eliminar usuarios", Clave = "USR_DELETE" },
                    new Permiso { Nombre = "Perfiles_Ver", Descripcion = "Ver perfiles", Clave = "PROF_VIEW" },
                    new Permiso { Nombre = "Perfiles_Crear", Descripcion = "Crear perfiles", Clave = "PROF_CREATE" },
                    new Permiso { Nombre = "Perfiles_Editar", Descripcion = "Editar perfiles", Clave = "PROF_EDIT" },
                    new Permiso { Nombre = "Perfiles_Eliminar", Descripcion = "Eliminar perfiles", Clave = "PROF_DELETE" }
                };

                await _context.Permisos.AddRangeAsync(permisos);
                await _context.SaveChangesAsync();

                // Crear perfil de Administrador con todos los permisos
                var perfilAdmin = new Perfil
                {
                    Nombre = "Administrador",
                    Descripcion = "Perfil con acceso total al sistema",
                    Permisos = await _context.Permisos.ToListAsync()
                };

                // Crear perfil de Usuario con permisos básicos
                var perfilUsuario = new Perfil
                {
                    Nombre = "Usuario",
                    Descripcion = "Perfil con acceso limitado",
                    Permisos = await _context.Permisos.Where(p => p.Nombre.StartsWith("Usuarios_Ver")).ToListAsync()
                };

                await _context.Perfiles.AddRangeAsync(new[] { perfilAdmin, perfilUsuario });
                await _context.SaveChangesAsync();
            }
        }

        private async Task CheckUsuariosAsync()
        {
            // Verificar si los usuarios admin o usuario ya existen
            var adminExists = await _context.Usuarios.AnyAsync(u => u.NombreUsuario == "admin" || u.Email == "admin@implanprod.com");
            var usuarioExists = await _context.Usuarios.AnyAsync(u => u.NombreUsuario == "usuario" || u.Email == "usuario@implanprod.com");

            // Si ambos usuarios existen, no hacer nada
            if (adminExists && usuarioExists)
            {
                return;
            }

            var perfilAdmin = await _context.Perfiles.FirstOrDefaultAsync(p => p.Nombre == "Administrador");
            var perfilUsuario = await _context.Perfiles.FirstOrDefaultAsync(p => p.Nombre == "Usuario");
            var empresa = await _context.Empresas.FirstOrDefaultAsync();
            var dependenciaAdmin = await _context.Dependencias.FirstOrDefaultAsync(d => d.dependencia == "Administración");
            var sectorRRHH = await _context.Sectores.FirstOrDefaultAsync(s => s.sector == "Recursos Humanos");
            var country = await _context.Countries.FirstOrDefaultAsync();
            var state = await _context.States.FirstOrDefaultAsync();
            var city = await _context.Cities.FirstOrDefaultAsync();

            if (perfilAdmin != null && perfilUsuario != null && empresa != null &&
                dependenciaAdmin != null && sectorRRHH != null && country != null &&
                state != null && city != null)
            {
                var usuarios = new List<Usuario>();

                // Agregar admin solo si no existe
                if (!adminExists)
                {
                    usuarios.Add(new Usuario
                    {
                        NombreUsuario = "admin",
                        Email = "admin@implanprod.com",
                        Password = HashPassword("Admin123!"),
                        FechaRegistro = DateTime.UtcNow,
                        Activo = true,
                        FotoUrl = null,
                        PerfilId = perfilAdmin.Id,
                        EmpresaId = empresa.Id,
                        DependenciaId = dependenciaAdmin.Id,
                        SectorId = sectorRRHH.Id,
                        CountryId = country.Id,
                        StateId = state.Id,
                        CityId = city.Id,
                        IntentosLogin = 0,
                        Bloqueado = false,
                        UltimoAcceso = DateTime.UtcNow,
                        RefreshToken = null,
                        RefreshTokenExpiryTime = null
                    });
                }

                // Agregar usuario solo si no existe
                if (!usuarioExists)
                {
                    usuarios.Add(new Usuario
                    {
                        NombreUsuario = "usuario",
                        Email = "usuario@implanprod.com",
                        Password = HashPassword("Usuario123!"),
                        FechaRegistro = DateTime.UtcNow,
                        Activo = true,
                        FotoUrl = null,
                        PerfilId = perfilUsuario.Id,
                        EmpresaId = empresa.Id,
                        DependenciaId = dependenciaAdmin.Id,
                        SectorId = sectorRRHH.Id,
                        CountryId = country.Id,
                        StateId = state.Id,
                        CityId = city.Id,
                        IntentosLogin = 0,
                        Bloqueado = false,
                        UltimoAcceso = DateTime.UtcNow,
                        RefreshToken = null,
                        RefreshTokenExpiryTime = null
                    });
                }

                // Solo agregar usuarios si hay alguno para agregar
                if (usuarios.Any())
                {
                    await _context.Usuarios.AddRangeAsync(usuarios);
                    await _context.SaveChangesAsync();
                }
            }
        }

        private string HashPassword(string password)
        {
            using (var sha256 = SHA256.Create())
            {
                var hashedBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                return Convert.ToBase64String(hashedBytes);
            }
        }

        //private async Task CheckUnidadMedidaAsync()
        //{
        //    if (!_context.UnidadMedida.Any())
        //    {
        //        _context.UnidadMedida.Add(new UnidadMedidas
        //        {
        //            DescripcionAbreviada = "LTS",
        //            Descripcion = "LITROS"
        //        });

        //        _context.UnidadMedida.Add(new UnidadMedidas
        //        {
        //            DescripcionAbreviada = "MTS",
        //            Descripcion = "METROS"
        //        });
        //        _context.UnidadMedida.Add(new UnidadMedidas
        //        {
        //            DescripcionAbreviada = "UN",
        //            Descripcion = "UNIDAD"
        //        });
        //        _context.UnidadMedida.Add(new UnidadMedidas
        //        {
        //            DescripcionAbreviada = "TN",
        //            Descripcion = "TONELADA"
        //        });
        //        _context.UnidadMedida.Add(new UnidadMedidas
        //        {
        //            DescripcionAbreviada = "PZ",
        //            Descripcion = "PIEZA"
        //        });
        //        _context.UnidadMedida.Add(new UnidadMedidas
        //        {
        //            DescripcionAbreviada = "KM",
        //            Descripcion = "KILOMETRO"
        //        });
        //        _context.UnidadMedida.Add(new UnidadMedidas
        //        {
        //            DescripcionAbreviada = "KG",
        //            Descripcion = "KILOGRAMO"
        //        });

        //        await _context.SaveChangesAsync();
        //    }
        //}

        private async Task CheckEmailConfigurationAsync()
        {
            if (!_context.Set<EmailConfiguration>().Any())
            {
                var defaultConfig = new EmailConfiguration
                {
                    EmailAddress = "sistema@example.com",
                    EncryptedPassword = "", // La contraseña debe ser encriptada antes de guardarla
                    DisplayName = "Sistema IMPLAN",
                    LastModified = DateTime.Now
                };

                await _context.Set<EmailConfiguration>().AddAsync(defaultConfig);
                await _context.SaveChangesAsync();
            }
        }
    }
}
//20/04/25
//OS COUNTRIES,STATES, CITIES HAY QUE CREARLE EL CONTENIDO SI NO NO ANDA CUANDO SE GENERA UNA BASE DE DATOS DE 0

//23/04/25
//ALTER TABLE proved12
//ALTER COLUMN prov_cli VARCHAR(48);

// TABLA TADEPOSI
//CREATE TABLE[dbo].[Tadeposi] (
//    [Id][int] IDENTITY(1, 1) NOT NULL,
//    [CodigoDepo] [int] NOT NULL,
//    [Descripcion] [nvarchar] (64) NOT NULL,
//    [DescripAbreviada] [nvarchar] (12) NOT NULL,
//    [Status] [int] NOT NULL DEFAULT 0,  -- Valor por defecto 0
//    [CodPlanta][int] NULL,
//    [Disponible][bit] NOT NULL DEFAULT 1,  -- Valor por defecto True (1 en SQL Server)
//    CONSTRAINT [PK_Tadeposi] PRIMARY KEY CLUSTERED 
//    (
//        [Id] ASC
//    ) WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF,
//            ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON[PRIMARY]
//) ON[PRIMARY];
//GO

// TABLA TIPO DE MATERIALES
//CREATE TABLE[dbo].[Tatimat] (
//    [Id][int] IDENTITY(1, 1) NOT NULL,
//    [CodigoTatimat] [int] NOT NULL,
//    [Descripcion] [nvarchar] (64) NOT NULL,
//    [Status] [int] NOT NULL,
//    [Deposito] [smallint] NOT NULL,
//    [Referencia] [nvarchar] (128)  NULL,
//    CONSTRAINT[PK_Tatimat] PRIMARY KEY CLUSTERED 
//    (
//        [Id] ASC
//    )WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF,
//          ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) 
//    ON[PRIMARY]
//) ON[PRIMARY]
//GO

//TABLA FAMIALIAS
//CREATE TABLE[dbo].[Familia] (
//    [Id][int] IDENTITY(1, 1) NOT NULL,

//    [CodigoFamilia] [nvarchar] (4) NOT NULL,

//    [Descripcion] [nvarchar] (64) NOT NULL,

//    [Status] [int] NOT NULL,
// CONSTRAINT[PK_Familia] PRIMARY KEY CLUSTERED 
//(

//    [Id] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON[PRIMARY]
//) ON[PRIMARY]
//GO

//tabla TAMONED
//CREATE TABLE[dbo].[Tamoned] (
//    [Id][int] IDENTITY(1, 1) NOT NULL,

//    [CodigoMoneda] [int] NOT NULL,

//    [Descripcion] [nvarchar] (32) NOT NULL,

//    [Status] [int] NOT NULL,

//    [Cotizacion] [decimal] (18, 4) NOT NULL,
// CONSTRAINT[PK_Tamoned] PRIMARY KEY CLUSTERED 
//(

//    [Id] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//) ON[PRIMARY]
//GO
//
//TABLA GRUPIVA
//CREATE TABLE GRUPIVA (
//    id INT IDENTITY(1,1) PRIMARY KEY,
//    CodigoID INT NOT NULL,
//    Referencia NVARCHAR(32) NOT NULL,
//    tasa DECIMAL(5,2) NOT NULL
//);

//AGREGAR LOS DATOS SI ESTA VACIA
//IF(SELECT COUNT(*) FROM GRUPIVA) = 0
//BEGIN
//    INSERT INTO GRUPIVA (CodigoID, Referencia, tasa)
//    VALUES 
//        (0, 'IVA 21 %', 21.00),
//        (1, 'IVA 10,5 %', 10.50),
//        (2, 'IVA 27 %', 27.00);
//END
//ALTER TABLE grupiva 
//ADD  Status INT NOT NULL DEFAULT 0;  agregar este cambio

//TABLA CATOPER
//CREATE TABLE[dbo].[Catoper] (
//    [Id][int] IDENTITY(1, 1) NOT NULL,

//    [CodigoCategoria] [int] NOT NULL,

//    [Descripcion] [nvarchar] (32) NOT NULL,

//    [Status] [int] NOT NULL DEFAULT 0,

//    [ValorHora] [decimal] (18, 4) NOT NULL DEFAULT 0,
// CONSTRAINT[PK_Catoper] PRIMARY KEY CLUSTERED 
//(

//    [Id] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//) ON[PRIMARY]
//GO

//CREATE TABLE[dbo].[Secprod] (
//    [Id][int] IDENTITY(1, 1) NOT NULL,

//    [CodigoSecProd] [int] NOT NULL,

//    [Descripcion] [nvarchar] (64) NOT NULL,

//    [Status] [int] NOT NULL,

//    [Deposito] [int] NOT NULL,
// CONSTRAINT[PK_Secprod] PRIMARY KEY CLUSTERED 
//(

//    [Id] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON[PRIMARY]
//) ON[PRIMARY]
//GO

//TABLA PADMAQ PADRON DE MAQUINAS
//CREATE TABLE[dbo].[PADMAQ] (
//    [Id][int] IDENTITY(1, 1) NOT NULL,

//    [CodigoPadMaq] [nvarchar] (16) NOT NULL,

//    [Descripcion] [nvarchar] (64) NOT NULL,

//    [Status] [int] NOT NULL,

//    [CodigoSecProd] [int] NOT NULL,

//    [CodigoPadMaqAlt] [nvarchar] (16)  NULL,

//    [ValorHora] [decimal] (18, 4) NOT NULL DEFAULT 0
// CONSTRAINT [PK_PADMAQ] PRIMARY KEY CLUSTERED 
//(

//    [Id] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//) ON[PRIMARY]
//GO

