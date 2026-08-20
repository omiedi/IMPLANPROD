using IMPLANPROD.Shared.DTOs;

namespace IMPLANPROD.Server.Services
{
    public interface ISeguridadRutasService
    {
        Task<int> CountSuperAdminsAsync(CancellationToken cancellationToken = default);
        Task<bool> IsCurrentUserSuperAdminAsync(CancellationToken cancellationToken = default);
        Task<List<SeguridadModuloDto>> GetModulosAsync(bool incluirConfiguracion, CancellationToken cancellationToken = default);
        Task<List<SeguridadPaginaDto>> GetPaginasByModuloAsync(int moduloId, bool incluirConfiguracion, CancellationToken cancellationToken = default);
        Task<List<int>> GetPaginasPermitidasByUsuarioAsync(int usuarioId, CancellationToken cancellationToken = default);
        Task AsignarPaginasUsuarioAsync(int usuarioId, List<int> paginaIds, CancellationToken cancellationToken = default);
        Task<bool> PuedeAccederRutaAsync(int usuarioId, string ruta, CancellationToken cancellationToken = default);
    }
}
