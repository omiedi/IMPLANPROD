using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Data.Interceptors;
using IMPLANPROD.Server.Middleware;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Server.Utilities;
using IMPLANPROD.Shared.Services;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.FileProviders;
using Microsoft.IdentityModel.Tokens;
using Serilog;
using Serilog.Events;
using System.Text;
using System.Text.Json.Serialization;

// TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Importación de namespaces necesarios para el servicio de contexto de usuario

var builder = WebApplication.CreateBuilder(args);

// Configurar Serilog
Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Debug()
    .MinimumLevel.Override("Microsoft", LogEventLevel.Information)
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .WriteTo.File("Logs/log-.txt", rollingInterval: RollingInterval.Day)
    .CreateLogger();

builder.Host.UseSerilog();

// Configuración de JWT
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidateAudience = true,
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true,
            ValidIssuer = builder.Configuration["Jwt:Issuer"],
            ValidAudience = builder.Configuration["Jwt:Audience"],
            IssuerSigningKey = new SymmetricSecurityKey(
                Encoding.UTF8.GetBytes(builder.Configuration["Jwt:Key"] ?? throw new InvalidOperationException("JWT Key no configurada")))
        };
    });

// Add services to the container.
builder.Services.AddControllersWithViews()
    .AddJsonOptions(options =>
    {
        options.JsonSerializerOptions.ReferenceHandler = ReferenceHandler.IgnoreCycles;
    });

builder.Services.AddRazorPages();

// Registrar el interceptor de auditoría como servicio
builder.Services.AddScoped<AuditInterceptor>();

// Configurar la base de datos con el interceptor de auditoría avanzado
builder.Services.AddDbContext<DataContext>((serviceProvider, options) =>
{
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"))
           .AddInterceptors(serviceProvider.GetRequiredService<AuditInterceptor>());
});

// Configurar servicios personalizados
builder.Services.AddTransient<SeedDb>();
builder.Services.AddScoped<IBlobStorageService, BlobStorageService>();
builder.Services.AddScoped<IEmailService, EmailService>();
builder.Services.AddScoped<DocumentNumberGenerator>();
builder.Services.AddScoped<IUsuarioFlexoftService, UsuarioFlexoftService>();//para obtener en los controladores el ususario flexoft
builder.Services.AddScoped<ICodiempresa, Codiempresa>();//para obtener en los controladores el codiempr formateado con dependencia
builder.Services.AddScoped<IMasterFuncionesService, MasterFuncionesService>();//funciones de tabla master las mas comunes
builder.Services.AddScoped<ExcelService>();//generico para exportar a excel
builder.Services.AddScoped<ReservasService>();//servicio para manejar las reservas de órdenes de fabricación
builder.Services.AddScoped<FlexcrlService>();//servicio para manejar controles de procesos de negocio (tabla FLEXCRL)
builder.Services.AddScoped<IMoviStockService, MoviStockService>();//servicio para manejar movimientos de stock (equivalente a MOVISTOK VB6)
builder.Services.AddScoped<FileStorageService>();//servicio para almacenamiento de archivos con ruta configurable por ambiente
builder.Services.AddScoped<AnotacionesPedidoService>();//servicio para manejar anotaciones automáticas de pedidos
builder.Services.AddScoped<ConversionMonedaService>();//servicio para manejar conversiones de moneda en pedidos
builder.Services.AddScoped<DetaivaService>();//servicio para manejar detalle de IVA en órdenes de compra (equivalente a DETAIVA VB6)
builder.Services.AddScoped<ConsignacionService>();//servicio para manejar consignaciones MACONSIG según lógica VB6
builder.Services.AddScoped<ActualizacionCostoMasterService>();//servicio para actualizar costos en MASTER según configuración ACOSREP (equivalente a lógica VB6 GRACOMPRA)
builder.Services.AddScoped<MigrationLegacyService>();//servicio para migración de datos legacy
builder.Services.AddScoped<IDerechosUsuarioService, DerechosUsuarioService>();//servicio para gestión de derechos de usuarios sobre aplicaciones
builder.Services.AddScoped<ISyncService, SyncService>();//servicio para sincronización Central-Sucursal
builder.Services.AddHostedService<SyncBackgroundService>();//servicio en segundo plano para procesamiento automático de sincronización (solo Sucursales)

// TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Registro del servicio de contexto de usuario y HttpContextAccessor
builder.Services.AddHttpContextAccessor(); // Necesario para acceder al contexto HTTP en el servicio
builder.Services.AddScoped<IUserContextService, UserContextService>(); // Servicio centralizado para información del usuario
builder.Services.AddScoped<IEmpresaDependenciaService, EmpresaDependenciaService>(); // Servicio para validaciones por empresa y dependencia

// Registrar el servicio Georef para datos geográficos de Argentina
builder.Services.AddHttpClient<GeorefService>();
builder.Services.AddScoped<GeorefService>();

// Configurar opciones de FileStorage desde appsettings.json
builder.Services.Configure<FileStorageOptions>(builder.Configuration.GetSection(FileStorageOptions.SectionName));

var app = builder.Build();

// Verifica si el entorno es de producción
IServiceScopeFactory? scopedFactory = app.Services.GetService<IServiceScopeFactory>();

using (IServiceScope? scope = scopedFactory!.CreateScope())
{
    SeedDb? service = scope.ServiceProvider.GetService<SeedDb>();
    service!.SeedAsync().Wait();
}

app.UseCors(c => c.AllowAnyHeader().AllowAnyOrigin().AllowAnyMethod());

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseWebAssemblyDebugging();
}
else
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseBlazorFrameworkFiles();

// Configurar archivos estáticos
app.UseStaticFiles(); // Para wwwroot

// Configurar archivos estáticos desde ruta externa (FileStorage)
var fileStorageOptions = builder.Configuration.GetSection(FileStorageOptions.SectionName).Get<FileStorageOptions>();
if (fileStorageOptions != null && fileStorageOptions.ServeStaticFiles && !string.IsNullOrEmpty(fileStorageOptions.BasePath))
{
    // Crear directorio si no existe
    if (!Directory.Exists(fileStorageOptions.BasePath))
    {
        Directory.CreateDirectory(fileStorageOptions.BasePath);
        Log.Information("Directorio de FileStorage creado: {BasePath}", fileStorageOptions.BasePath);
    }

    app.UseStaticFiles(new StaticFileOptions
    {
        FileProvider = new PhysicalFileProvider(fileStorageOptions.BasePath),
        RequestPath = "/uploads"
    });
    
    Log.Information("Archivos estáticos configurados desde: {BasePath}", fileStorageOptions.BasePath);
}
else
{
    Log.Warning("FileStorage no configurado, usando ruta por defecto en wwwroot");
}

// IMPORTANTE: UseRouting debe ir ANTES de UseAuthentication y UseAuthorization
app.UseRouting();

// Agregar autenticación y autorización
app.UseAuthentication();

// TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Registro del middleware de contexto de usuario
// Se coloca después de UseAuthentication y antes de UseAuthorization para tener acceso al usuario autenticado
app.UseUserContext();

app.UseAuthorization();

app.MapRazorPages();
app.MapControllers();
app.MapFallbackToFile("index.html");

app.Run();
