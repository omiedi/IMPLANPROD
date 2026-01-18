using IMPLANPROD.Client.Repositories;

namespace IMPLANPROD.Client.Services
{
    /// <summary>
    /// Servicio para obtener y validar la dependencia activa del usuario
    /// </summary>
    public class DependenciaService
    {
        private readonly IRepository _repository;
        private InfoUsuarioDTO? _infoUsuarioCache;
        private DateTime? _ultimaActualizacion;
        private readonly TimeSpan _tiempoCache = TimeSpan.FromMinutes(5);

        public DependenciaService(IRepository repository)
        {
            _repository = repository;
        }

        /// <summary>
        /// Obtiene la información del usuario actual (con caché)
        /// </summary>
        private async Task<InfoUsuarioDTO?> ObtenerInfoUsuarioAsync()
        {
            // Si hay caché válido, retornarlo
            if (_infoUsuarioCache != null && 
                _ultimaActualizacion.HasValue && 
                DateTime.Now - _ultimaActualizacion.Value < _tiempoCache)
            {
                return _infoUsuarioCache;
            }

            // Obtener información fresca del servidor
            try
            {
                var response = await _repository.Get<InfoUsuarioDTO>("api/Usuarios/InfoUsuarioActual");
                
                if (!response.Error && response.Response != null)
                {
                    _infoUsuarioCache = response.Response;
                    _ultimaActualizacion = DateTime.Now;
                    return _infoUsuarioCache;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener información del usuario: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Obtiene la dependencia activa del usuario
        /// </summary>
        /// <returns>Nombre de la dependencia o null si no se pudo obtener</returns>
        public async Task<string?> ObtenerDependenciaActivaAsync()
        {
            var info = await ObtenerInfoUsuarioAsync();
            return info?.Dependencia;
        }

        /// <summary>
        /// Obtiene la empresa activa del usuario
        /// </summary>
        /// <returns>Nombre de la empresa o null si no se pudo obtener</returns>
        public async Task<string?> ObtenerEmpresaActivaAsync()
        {
            var info = await ObtenerInfoUsuarioAsync();
            return info?.Empresa;
        }

        /// <summary>
        /// Valida si la dependencia activa coincide con el nombre especificado (case-insensitive)
        /// </summary>
        /// <param name="nombreDependencia">Nombre de la dependencia a validar</param>
        /// <returns>True si coincide, False si no coincide</returns>
        public async Task<bool> EsDependenciaAsync(string nombreDependencia)
        {
            if (string.IsNullOrWhiteSpace(nombreDependencia))
                return false;

            var dependenciaActual = await ObtenerDependenciaActivaAsync();
            
            if (string.IsNullOrWhiteSpace(dependenciaActual))
                return false;

            return dependenciaActual.Equals(nombreDependencia, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Valida si la dependencia activa está en la lista de dependencias especificadas (case-insensitive)
        /// </summary>
        /// <param name="dependencias">Lista de nombres de dependencias a validar</param>
        /// <returns>True si la dependencia activa está en la lista, False si no</returns>
        public async Task<bool> EsAlgunaDependenciaAsync(params string[] dependencias)
        {
            if (dependencias == null || dependencias.Length == 0)
                return false;

            var dependenciaActual = await ObtenerDependenciaActivaAsync();
            
            if (string.IsNullOrWhiteSpace(dependenciaActual))
                return false;

            return dependencias.Any(d => 
                !string.IsNullOrWhiteSpace(d) && 
                dependenciaActual.Equals(d, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Valida si la empresa activa coincide con el nombre especificado (case-insensitive)
        /// </summary>
        /// <param name="nombreEmpresa">Nombre de la empresa a validar</param>
        /// <returns>True si coincide, False si no coincide</returns>
        public async Task<bool> EsEmpresaAsync(string nombreEmpresa)
        {
            if (string.IsNullOrWhiteSpace(nombreEmpresa))
                return false;

            var empresaActual = await ObtenerEmpresaActivaAsync();
            
            if (string.IsNullOrWhiteSpace(empresaActual))
                return false;

            return empresaActual.Equals(nombreEmpresa, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Limpia el caché forzando una nueva consulta en la próxima llamada
        /// </summary>
        public void LimpiarCache()
        {
            _infoUsuarioCache = null;
            _ultimaActualizacion = null;
        }

        /// <summary>
        /// DTO para recibir la información del usuario
        /// </summary>
        private class InfoUsuarioDTO
        {
            public int UsuarioId { get; set; }
            public int IdFlexoft { get; set; }
            public string Empresa { get; set; } = "";
            public string Dependencia { get; set; } = "";
        }
    }
}
