using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    public class SeguridadRutasService : ISeguridadRutasService
    {
        private readonly DataContext _context;
        private readonly IUserContextService _userContextService;

        public SeguridadRutasService(DataContext context, IUserContextService userContextService)
        {
            _context = context;
            _userContextService = userContextService;
        }

        public Task<int> CountSuperAdminsAsync(CancellationToken cancellationToken = default)
        {
            return _context.Usuarios
                .AsNoTracking()
                .Include(x => x.Perfil)
                .CountAsync(x => x.Activo && x.Perfil != null && EsSuperAdmin(x.Perfil.Nombre), cancellationToken);
        }

        public async Task<bool> IsCurrentUserSuperAdminAsync(CancellationToken cancellationToken = default)
        {
            var currentUserId = _userContextService.GetCurrentUsuarioId();
            if (!currentUserId.HasValue)
            {
                return false;
            }

            var perfilNombre = await _context.Usuarios
                .AsNoTracking()
                .Where(x => x.Id == currentUserId.Value)
                .Select(x => x.Perfil != null ? x.Perfil.Nombre : null)
                .FirstOrDefaultAsync(cancellationToken);

            return EsSuperAdmin(perfilNombre);
        }

        public async Task<List<SeguridadModuloDto>> GetModulosAsync(bool incluirConfiguracion, CancellationToken cancellationToken = default)
        {
            var query = _context.SegModulos
                .AsNoTracking()
                .Where(x => x.Activo)
                .OrderBy(x => x.Orden)
                .ThenBy(x => x.Nombre)
                .Select(x => new SeguridadModuloDto
                {
                    Id = x.Id,
                    Codigo = x.Codigo,
                    Nombre = x.Nombre,
                    Orden = x.Orden,
                    Paginas = x.Paginas
                        .Where(p => p.Activo && (incluirConfiguracion || !p.EsConfiguracion))
                        .OrderBy(p => p.Orden)
                        .ThenBy(p => p.Nombre)
                        .Select(p => new SeguridadPaginaDto
                        {
                            Id = p.Id,
                            ModuloId = p.ModuloId,
                            Nombre = p.Nombre,
                            Ruta = p.Ruta,
                            EsConfiguracion = p.EsConfiguracion,
                            Orden = p.Orden
                        })
                        .ToList()
                });

            return await query.ToListAsync(cancellationToken);
        }

        public Task<List<SeguridadPaginaDto>> GetPaginasByModuloAsync(int moduloId, bool incluirConfiguracion, CancellationToken cancellationToken = default)
        {
            return _context.SegPaginas
                .AsNoTracking()
                .Where(x => x.ModuloId == moduloId && x.Activo && (incluirConfiguracion || !x.EsConfiguracion))
                .OrderBy(x => x.Orden)
                .ThenBy(x => x.Nombre)
                .Select(x => new SeguridadPaginaDto
                {
                    Id = x.Id,
                    ModuloId = x.ModuloId,
                    Nombre = x.Nombre,
                    Ruta = x.Ruta,
                    EsConfiguracion = x.EsConfiguracion,
                    Orden = x.Orden
                })
                .ToListAsync(cancellationToken);
        }

        public Task<List<int>> GetPaginasPermitidasByUsuarioAsync(int usuarioId, CancellationToken cancellationToken = default)
        {
            return _context.SegUsuarioPaginas
                .AsNoTracking()
                .Where(x => x.UsuarioId == usuarioId && x.Permitido)
                .Select(x => x.PaginaId)
                .ToListAsync(cancellationToken);
        }

        public async Task AsignarPaginasUsuarioAsync(int usuarioId, List<int> paginaIds, CancellationToken cancellationToken = default)
        {
            var usuario = await _context.Usuarios
                .FirstOrDefaultAsync(x => x.Id == usuarioId, cancellationToken);

            if (usuario == null)
            {
                throw new InvalidOperationException("Usuario no encontrado.");
            }

            var paginaIdsNormalizados = (paginaIds ?? new List<int>())
                .Where(x => x > 0)
                .Distinct()
                .ToList();

            if (paginaIdsNormalizados.Count > 0)
            {
                var paginasValidasCount = await _context.SegPaginas
                    .AsNoTracking()
                    .CountAsync(x => paginaIdsNormalizados.Contains(x.Id) && x.Activo, cancellationToken);

                if (paginasValidasCount != paginaIdsNormalizados.Count)
                {
                    throw new InvalidOperationException("Se detectaron páginas inválidas o inactivas en la asignación.");
                }
            }

            var existentes = await _context.SegUsuarioPaginas
                .Where(x => x.UsuarioId == usuarioId)
                .ToListAsync(cancellationToken);

            _context.SegUsuarioPaginas.RemoveRange(existentes);

            foreach (var paginaId in paginaIdsNormalizados)
            {
                _context.SegUsuarioPaginas.Add(new Shared.Entities.SegUsuarioPagina
                {
                    UsuarioId = usuarioId,
                    PaginaId = paginaId,
                    Permitido = true,
                    Codiempr = 0
                });
            }

            await _context.SaveChangesAsync(cancellationToken);
        }

        public async Task<bool> PuedeAccederRutaAsync(int usuarioId, string ruta, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(ruta))
            {
                return false;
            }

            var usuario = await _context.Usuarios
                .AsNoTracking()
                .Include(x => x.Perfil)
                .FirstOrDefaultAsync(x => x.Id == usuarioId && x.Activo, cancellationToken);

            if (usuario == null)
            {
                return false;
            }

            var perfilNombre = usuario.Perfil?.Nombre;
            if (EsSuperAdmin(perfilNombre))
            {
                return true;
            }

            // Normalizar: quitar query string y slash final para comparación limpia
            // Ej: "/cities/edit/5?tab=1" → "/cities/edit/5"
            var rutaNorm = (ruta.Split('?')[0]).TrimEnd('/').ToLowerInvariant();

            // Prefix match: busca la página registrada cuya Ruta sea prefijo de la URL real.
            // Ej: URL="/cities/edit/5" → matchea con Ruta="/cities/edit" (la más larga/específica gana).
            // Se carga el conjunto reducido de páginas activas para hacer el match en memoria.
            var todasLasPaginas = await _context.SegPaginas
                .AsNoTracking()
                .Where(x => x.Activo)
                .Select(x => new { x.Id, x.Ruta, x.EsConfiguracion })
                .ToListAsync(cancellationToken);

            var pagina = todasLasPaginas
                .Where(x => rutaNorm.StartsWith(x.Ruta.TrimEnd('/').ToLowerInvariant(), StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(x => x.Ruta.Length)
                .FirstOrDefault();

            if (pagina == null)
            {
                return false;
            }

            if (EsAdministrador(perfilNombre))
            {
                return !pagina.EsConfiguracion;
            }

            return await _context.SegUsuarioPaginas
                .AsNoTracking()
                .AnyAsync(x => x.UsuarioId == usuarioId && x.PaginaId == pagina.Id && x.Permitido, cancellationToken);
        }

        private static bool EsSuperAdmin(string? perfilNombre)
        {
            return string.Equals((perfilNombre ?? string.Empty).Trim(), "SUPERADMIN", StringComparison.OrdinalIgnoreCase);
        }

        private static bool EsAdministrador(string? perfilNombre)
        {
            return string.Equals((perfilNombre ?? string.Empty).Trim(), "ADMINISTRADOR", StringComparison.OrdinalIgnoreCase)
                || string.Equals((perfilNombre ?? string.Empty).Trim(), "ADMIN", StringComparison.OrdinalIgnoreCase);
        }
    }
}
