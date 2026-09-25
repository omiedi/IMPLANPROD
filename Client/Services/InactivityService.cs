using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using CurrieTechnologies.Razor.SweetAlert2;
using IMPLANPROD.Client.Auth;

namespace IMPLANPROD.Client.Services
{
    /// <summary>
    /// Servicio que monitorea la inactividad del usuario y cierra la sesión automáticamente
    /// después de un período configurable de tiempo sin actividad.
    ///
    /// NOTA (2025-06-24): Este servicio fue desvinculado del DI y de MainLayout porque
    /// la sesión ahora depende únicamente de la expiración del token JWT del servidor.
    /// Se mantiene en el repositorio como referencia histórica; para reactivarlo hay que
    /// volver a registrarlo en Program.cs y conectarlo en MainLayout.
    /// </summary>
    public class InactivityService : IDisposable
    {
        private readonly NavigationManager _navigationManager;
        private readonly AuthenticationService _authenticationService;
        private readonly SweetAlertService _sweetAlertService;
        private Timer? _inactivityTimer;
        private DateTime _lastActivityTime;
        private bool _isWarningShown = false;
        private bool _disposed = false;

        /// <summary>
        /// Tiempo de inactividad en minutos antes de cerrar sesión (configurable)
        /// Por defecto: 20 minutos
        /// </summary>
        public int InactivityTimeoutMinutes { get; set; } = 20;

        /// <summary>
        /// Tiempo de advertencia en minutos antes del cierre (muestra alerta)
        /// Por defecto: 2 minutos antes del cierre
        /// </summary>
        public int WarningTimeoutMinutes { get; set; } = 2;

        /// <summary>
        /// Indica si el monitoreo de inactividad está activo
        /// </summary>
        public bool IsMonitoring { get; private set; }

        public InactivityService(
            NavigationManager navigationManager,
            AuthenticationStateProvider authenticationStateProvider,
            SweetAlertService sweetAlertService)
        {
            _navigationManager = navigationManager;
            _authenticationService = (AuthenticationService)authenticationStateProvider;
            _sweetAlertService = sweetAlertService;
            _lastActivityTime = DateTime.Now;
        }

        /// <summary>
        /// Inicia el monitoreo de inactividad
        /// </summary>
        public void StartMonitoring()
        {
            if (IsMonitoring) return;

            Console.WriteLine($"[InactivityService] Iniciando monitoreo - Timeout: {InactivityTimeoutMinutes} minutos");
            
            _lastActivityTime = DateTime.Now;
            _isWarningShown = false;
            IsMonitoring = true;

            // Timer que verifica cada 30 segundos
            _inactivityTimer = new Timer(CheckInactivity, null, TimeSpan.Zero, TimeSpan.FromSeconds(30));
        }

        /// <summary>
        /// Detiene el monitoreo de inactividad
        /// </summary>
        public void StopMonitoring()
        {
            if (!IsMonitoring) return;

            Console.WriteLine("[InactivityService] Deteniendo monitoreo");
            
            _inactivityTimer?.Dispose();
            _inactivityTimer = null;
            IsMonitoring = false;
            _isWarningShown = false;
        }

        /// <summary>
        /// Registra actividad del usuario (debe llamarse desde eventos de mouse/teclado)
        /// </summary>
        public void RegisterActivity()
        {
            _lastActivityTime = DateTime.Now;
            
            // Si se mostró la advertencia y el usuario vuelve a estar activo, resetear
            if (_isWarningShown)
            {
                Console.WriteLine("[InactivityService] Usuario activo nuevamente - Reseteando advertencia");
                _isWarningShown = false;
            }
        }

        /// <summary>
        /// Verifica el tiempo de inactividad y toma acciones según corresponda
        /// </summary>
        private async void CheckInactivity(object? state)
        {
            if (!IsMonitoring) return;

            var inactiveTime = DateTime.Now - _lastActivityTime;
            var inactiveMinutes = inactiveTime.TotalMinutes;

            Console.WriteLine($"[InactivityService] Tiempo inactivo: {inactiveMinutes:F1} minutos");

            // Si superó el tiempo de inactividad, cerrar sesión
            if (inactiveMinutes >= InactivityTimeoutMinutes)
            {
                Console.WriteLine("[InactivityService] ⏰ TIMEOUT - Cerrando sesión por inactividad");
                await LogoutDueToInactivity();
            }
            // Si está cerca del timeout, mostrar advertencia
            else if (inactiveMinutes >= (InactivityTimeoutMinutes - WarningTimeoutMinutes) && !_isWarningShown)
            {
                Console.WriteLine("[InactivityService] ⚠️ Mostrando advertencia de inactividad");
                await ShowInactivityWarning();
            }
        }

        /// <summary>
        /// Muestra una advertencia al usuario indicando que la sesión está por expirar
        /// </summary>
        private async Task ShowInactivityWarning()
        {
            _isWarningShown = true;

            var remainingMinutes = InactivityTimeoutMinutes - (DateTime.Now - _lastActivityTime).TotalMinutes;

            await _sweetAlertService.FireAsync(new SweetAlertOptions
            {
                Title = "Sesión por expirar",
                Text = $"Su sesión expirará en {remainingMinutes:F0} minutos por inactividad. Mueva el mouse o presione una tecla para continuar.",
                Icon = SweetAlertIcon.Warning,
                ShowCancelButton = false,
                ConfirmButtonText = "Entendido",
                Timer = 10000 // Se cierra automáticamente en 10 segundos
            });

            // Registrar actividad al cerrar la alerta (el usuario interactuó)
            RegisterActivity();
        }

        /// <summary>
        /// Cierra la sesión del usuario por inactividad y redirige al login
        /// </summary>
        private async Task LogoutDueToInactivity()
        {
            StopMonitoring();

            await _sweetAlertService.FireAsync(new SweetAlertOptions
            {
                Title = "Sesión cerrada",
                Text = "Su sesión ha sido cerrada por inactividad. Por favor, inicie sesión nuevamente.",
                Icon = SweetAlertIcon.Info,
                ShowCancelButton = false,
                ConfirmButtonText = "Ir a Login",
                AllowOutsideClick = false,
                AllowEscapeKey = false
            });

            // Cerrar sesión
            await _authenticationService.LogoutAsync();

            // Redirigir a login
            _navigationManager.NavigateTo("/login", true);
        }

        /// <summary>
        /// Libera los recursos del servicio
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;

            StopMonitoring();
            _disposed = true;
        }
    }
}
