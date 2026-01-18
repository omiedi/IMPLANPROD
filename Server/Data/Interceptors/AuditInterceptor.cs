using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using System.Security.Claims;

namespace IMPLANPROD.Server.Data.Interceptors
{
    /// <summary>
    /// Interceptor avanzado que maneja auditoría completa:
    /// - Timestamps automáticos (AddRecord, LastUpdate)
    /// - Información de usuario y empresa (NumUsuar, CodiEmprNet)
    /// </summary>
    public class AuditInterceptor : SaveChangesInterceptor
    {
        private readonly IServiceProvider _serviceProvider;

        public AuditInterceptor(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        /// <summary>
        /// Se ejecuta antes de guardar cambios en la base de datos (versión síncrona)
        /// </summary>
        public override InterceptionResult<int> SavingChanges(DbContextEventData eventData, InterceptionResult<int> result)
        {
            UpdateAuditFields(eventData.Context);
            return base.SavingChanges(eventData, result);
        }

        /// <summary>
        /// Se ejecuta antes de guardar cambios en la base de datos (versión asíncrona)
        /// </summary>
        public override ValueTask<InterceptionResult<int>> SavingChangesAsync(DbContextEventData eventData, InterceptionResult<int> result, CancellationToken cancellationToken = default)
        {
            UpdateAuditFields(eventData.Context);
            return base.SavingChangesAsync(eventData, result, cancellationToken);
        }

        /// <summary>
        /// Actualiza los campos de auditoría para todas las entidades modificadas
        /// Valida que se pueda obtener información de usuario y empresa antes de proceder
        /// </summary>
        private void UpdateAuditFields(DbContext? context)
        {
            if (context == null) return;

            var now = DateTime.Now;
            int? currentUserId = null;
            short? currentCompanyCode = null;
            bool hasValidSession = false;

            // Obtener información del usuario actual usando los servicios inyectados
            try
            {
                using var scope = _serviceProvider.CreateScope();
                var usuarioService = scope.ServiceProvider.GetService<IUsuarioFlexoftService>();
                var empresaService = scope.ServiceProvider.GetService<ICodiempresa>();
                var httpContextAccessor = scope.ServiceProvider.GetService<IHttpContextAccessor>();

                if (httpContextAccessor?.HttpContext?.User != null)
                {
                    var user = httpContextAccessor.HttpContext.User;

                    // Verificar que el usuario esté autenticado
                    if (user.Identity?.IsAuthenticated == true)
                    {
                        // Obtener ID del usuario
                        currentUserId = usuarioService?.InterObtenerIdFlexoft(user);

                        // Obtener código de empresa
                        var empresaCode = empresaService?.IObtenerCodiempresa(user);
                        if (empresaCode.HasValue)
                        {
                            currentCompanyCode = empresaCode.Value;
                        }

                        // Validar que se obtuvieron ambos valores críticos
                        hasValidSession = IsValidSessionData(currentUserId, currentCompanyCode);

                        // Logging para debugging
                        if (hasValidSession)
                        {
                            Console.WriteLine($"[AuditInterceptor] Sesión válida - Usuario: {currentUserId}, Empresa: {currentCompanyCode}");
                        }
                        else
                        {
                            Console.WriteLine($"[AuditInterceptor] Sesión inválida - Usuario: {currentUserId}, Empresa: {currentCompanyCode}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error obteniendo información de usuario/empresa: {ex.Message}");
                hasValidSession = false;
            }

            // Verificar si hay entidades que requieren auditoría
            var auditableEntries = context.ChangeTracker.Entries()
                .Where(e => e.Entity is IAuditableEntity &&
                            (e.State == EntityState.Added || e.State == EntityState.Modified))
                .ToList();

            // Excluir entidad Usuario de la validación estricta (necesaria para el proceso de login)
            var entriesRequiringValidation = auditableEntries
                .Where(e => e.Entity.GetType().Name != "Usuario")
                .ToList();

            // Si hay entidades (excepto Usuario) que requieren auditoría pero no hay sesión válida, lanzar excepción
            if (entriesRequiringValidation.Any() && !hasValidSession)
            {
                var entityNames = entriesRequiringValidation
                    .Select(e => e.Entity.GetType().Name)
                    .Distinct()
                    .ToList();

                // Logging detallado del problema antes de lanzar la excepción
                Console.WriteLine($"[AuditInterceptor] ❌ BLOQUEO DE SEGURIDAD ACTIVADO");
                Console.WriteLine($"[AuditInterceptor] Motivo: Usuario no logueado o datos de sesión inválidos");
                Console.WriteLine($"[AuditInterceptor] Usuario ID: {(currentUserId?.ToString() ?? "NO DISPONIBLE")}");
                Console.WriteLine($"[AuditInterceptor] Código Empresa: {(currentCompanyCode?.ToString() ?? "NO DISPONIBLE")}");
                Console.WriteLine($"[AuditInterceptor] Entidades afectadas: {string.Join(", ", entityNames)}");
                Console.WriteLine($"[AuditInterceptor] Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

                throw CreateAuditException(entityNames, currentUserId, currentCompanyCode);
            }

            // Procesar entidades con auditoría completa (IAuditableEntity)
            // Usar la lista ya obtenida anteriormente
            foreach (var entry in auditableEntries)
            {
                var auditableEntity = (IAuditableEntity)entry.Entity;
                var isUsuarioEntity = entry.Entity.GetType().Name == "Usuario";

                if (entry.State == EntityState.Added)
                {
                    // Nuevo registro: establecer todos los campos de auditoría
                    auditableEntity.AddRecord = now;
                    auditableEntity.LastUpdate = now;
                    
                    // Solo asignar usuario/empresa si hay sesión válida O si no es la entidad Usuario
                    if (hasValidSession || !isUsuarioEntity)
                    {
                        auditableEntity.NumUsuar = currentUserId;
                        auditableEntity.CodiEmprNet = currentCompanyCode;
                    }
                }
                else if (entry.State == EntityState.Modified)
                {
                    // Registro modificado: solo actualizar LastUpdate y información actual
                    auditableEntity.LastUpdate = now;
                    
                    // Solo asignar usuario/empresa si hay sesión válida O si no es la entidad Usuario
                    if (hasValidSession || !isUsuarioEntity)
                    {
                        auditableEntity.NumUsuar = currentUserId;
                        auditableEntity.CodiEmprNet = currentCompanyCode;
                    }

                    // Preservar AddRecord original (no debe cambiar)
                    entry.Property(nameof(IAuditableEntity.AddRecord)).IsModified = false;
                }
            }

            // Procesar entidades con solo timestamps (ITimestampEntity pero no IAuditableEntity)
            var timestampOnlyEntries = context.ChangeTracker.Entries()
                .Where(e => e.Entity is ITimestampEntity &&
                            !(e.Entity is IAuditableEntity) &&
                            (e.State == EntityState.Added || e.State == EntityState.Modified));

            foreach (var entry in timestampOnlyEntries)
            {
                var timestampEntity = (ITimestampEntity)entry.Entity;

                if (entry.State == EntityState.Added)
                {
                    timestampEntity.AddRecord = now;
                    timestampEntity.LastUpdate = now;
                }
                else if (entry.State == EntityState.Modified)
                {
                    timestampEntity.LastUpdate = now;
                    // Preservar AddRecord original
                    entry.Property(nameof(ITimestampEntity.AddRecord)).IsModified = false;
                }
            }
        }

        /// <summary>
        /// Valida que la información de sesión sea válida para operaciones de auditoría
        /// </summary>
        private static bool IsValidSessionData(int? userId, short? companyCode)
        {
            return userId.HasValue && userId.Value > 0 && 
                   companyCode.HasValue && companyCode.Value > 0;
        }

        /// <summary>
        /// Crea una excepción detallada cuando no se puede obtener información de auditoría
        /// </summary>
        private static UnauthorizedAccessException CreateAuditException(
            List<string> entityNames, 
            int? userId, 
            short? companyCode)
        {
            var userInfo = userId.HasValue ? userId.ToString() : "NO DISPONIBLE";
            var companyInfo = companyCode.HasValue ? companyCode.ToString() : "NO DISPONIBLE";
            
            // Determinar el motivo específico del fallo
            var motivoEspecifico = "";
            if (!userId.HasValue || userId.Value <= 0)
            {
                motivoEspecifico = "El usuario no está logueado o el ID de usuario es inválido.";
            }
            else if (!companyCode.HasValue || companyCode.Value <= 0)
            {
                motivoEspecifico = "No se pudo obtener el código de empresa del usuario logueado.";
            }
            else
            {
                motivoEspecifico = "Los datos de sesión no cumplen con los requisitos de validación.";
            }
            
            return new UnauthorizedAccessException(
                $"❌ OPERACIÓN BLOQUEADA POR SEGURIDAD DE AUDITORÍA\n\n" +
                $"🔍 DIAGNÓSTICO:\n" +
                $"   • Motivo: {motivoEspecifico}\n" +
                $"   • Usuario ID: {userInfo}\n" +
                $"   • Código Empresa: {companyInfo}\n\n" +
                $"📋 ENTIDADES AFECTADAS:\n" +
                $"   • {string.Join("\n   • ", entityNames)}\n\n" +
                $"🔧 SOLUCIÓN:\n" +
                $"   1. Verifique que esté correctamente logueado en el sistema\n" +
                $"   2. Si el problema persiste, cierre sesión y vuelva a iniciar sesión\n" +
                $"   3. Contacte al administrador si el error continúa\n\n" +
                $"ℹ️  INFORMACIÓN TÉCNICA:\n" +
                $"   Esta validación es obligatoria para mantener la trazabilidad\n" +
                $"   y auditoría de todas las operaciones en la base de datos.\n" +
                $"   Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        }
    }
}
