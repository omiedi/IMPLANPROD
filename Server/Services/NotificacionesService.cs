using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.DTOs.Notificaciones;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    public class NotificacionesService
    {
        private readonly DataContext _context;
        private readonly IUserContextService _userContext;

        public NotificacionesService(DataContext context, IUserContextService userContext)
        {
            _context = context;
            _userContext = userContext;
        }

        public int GetCurrentUsuarioIdOrThrow()
        {
            var usuarioId = _userContext.GetCurrentUsuarioId();
            if (!usuarioId.HasValue || usuarioId.Value <= 0)
            {
                throw new UnauthorizedAccessException("Usuario no autenticado");
            }

            return usuarioId.Value;
        }

        public async Task<int> CrearAsync(CrearNotificacionRequestDTO request, CancellationToken cancellationToken = default)
        {
            var creadorUsuarioId = GetCurrentUsuarioIdOrThrow();

            var usuarioIds = (request.UsuarioIds ?? new List<int>())
                .Where(x => x > 0)
                .Distinct()
                .ToList();

            if (!request.EsParaTodos && usuarioIds.Count == 0)
            {
                throw new InvalidOperationException("Debe indicar destinatarios (UsuarioIds) o marcar EsParaTodos");
            }

            var notificacion = new Notificacion
            {
                Titulo = request.Titulo,
                Resumen = request.Resumen,
                DetalleHtml = request.DetalleHtml,
                FechaExpiracion = request.FechaExpiracion,
                EsParaTodos = request.EsParaTodos,
                CreadorUsuarioId = creadorUsuarioId,
                EsGeneradaPorRegla = false,
                ReglaSqlId = null
            };

            _context.Notificaciones.Add(notificacion);
            await _context.SaveChangesAsync(cancellationToken);

            if (!request.EsParaTodos)
            {
                var destinatarios = usuarioIds.Select(uid => new NotificacionDestinatario
                {
                    NotificacionId = notificacion.Id,
                    UsuarioId = uid
                });

                _context.NotificacionDestinatarios.AddRange(destinatarios);
                await _context.SaveChangesAsync(cancellationToken);
            }

            return notificacion.Id;
        }

        public async Task<List<NotificacionItemDTO>> ListarParaUsuarioActualAsync(PaginationDTO pagination, CancellationToken cancellationToken = default)
        {
            var usuarioId = GetCurrentUsuarioIdOrThrow();
            var now = DateTime.Now;

            var queryable = _context.Notificaciones
                .AsNoTracking()
                .Where(n => !n.FechaExpiracion.HasValue || n.FechaExpiracion > now)
                .Where(n =>
                    n.EsParaTodos ||
                    _context.NotificacionDestinatarios.Any(d => d.NotificacionId == n.Id && d.UsuarioId == usuarioId))
                .Where(n => !_context.NotificacionOcultas.Any(o => o.NotificacionId == n.Id && o.UsuarioId == usuarioId));

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                var filter = pagination.Filter.ToLower();
                queryable = queryable.Where(n =>
                    n.Titulo.ToLower().Contains(filter) ||
                    (n.Resumen != null && n.Resumen.ToLower().Contains(filter)));
            }

            return await queryable
                .OrderByDescending(n => n.AddRecord)
                .Paginate(pagination)
                .Select(n => new NotificacionItemDTO
                {
                    Id = n.Id,
                    Titulo = n.Titulo,
                    Resumen = n.Resumen,
                    AddRecord = n.AddRecord,
                    Leida = _context.NotificacionLeidas.Any(l => l.NotificacionId == n.Id && l.UsuarioId == usuarioId),
                    TieneDetalle = n.DetalleHtml != null && n.DetalleHtml != "",
                    EsParaTodos = n.EsParaTodos,
                    CreadorUsuarioId = n.CreadorUsuarioId,
                    EsGeneradaPorRegla = n.EsGeneradaPorRegla,
                    ReglaSqlId = n.ReglaSqlId,
                    DestinatariosResumen = n.EsParaTodos
                        ? "Todos"
                        : ("Usuarios: " + _context.NotificacionDestinatarios.Count(d => d.NotificacionId == n.Id))
                })
                .ToListAsync(cancellationToken);
        }

        public async Task<double> TotalPagesParaUsuarioActualAsync(PaginationDTO pagination, CancellationToken cancellationToken = default)
        {
            var usuarioId = GetCurrentUsuarioIdOrThrow();
            var now = DateTime.Now;

            var queryable = _context.Notificaciones
                .AsNoTracking()
                .Where(n => !n.FechaExpiracion.HasValue || n.FechaExpiracion > now)
                .Where(n =>
                    n.EsParaTodos ||
                    _context.NotificacionDestinatarios.Any(d => d.NotificacionId == n.Id && d.UsuarioId == usuarioId))
                .Where(n => !_context.NotificacionOcultas.Any(o => o.NotificacionId == n.Id && o.UsuarioId == usuarioId));

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                var filter = pagination.Filter.ToLower();
                queryable = queryable.Where(n =>
                    n.Titulo.ToLower().Contains(filter) ||
                    (n.Resumen != null && n.Resumen.ToLower().Contains(filter)));
            }

            double count = await queryable.CountAsync(cancellationToken);
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return totalPages;
        }

        public async Task<int> ContarNoLeidasAsync(CancellationToken cancellationToken = default)
        {
            var usuarioId = GetCurrentUsuarioIdOrThrow();
            var now = DateTime.Now;

            var queryable = _context.Notificaciones
                .AsNoTracking()
                .Where(n => !n.FechaExpiracion.HasValue || n.FechaExpiracion > now)
                .Where(n =>
                    n.EsParaTodos ||
                    _context.NotificacionDestinatarios.Any(d => d.NotificacionId == n.Id && d.UsuarioId == usuarioId))
                .Where(n => !_context.NotificacionOcultas.Any(o => o.NotificacionId == n.Id && o.UsuarioId == usuarioId))
                .Where(n => !_context.NotificacionLeidas.Any(l => l.NotificacionId == n.Id && l.UsuarioId == usuarioId));

            return await queryable.CountAsync(cancellationToken);
        }

        public async Task MarcarLeidaAsync(int notificacionId, CancellationToken cancellationToken = default)
        {
            var usuarioId = GetCurrentUsuarioIdOrThrow();
            var now = DateTime.Now;

            var hasAccess = await _context.Notificaciones
                .AsNoTracking()
                .AnyAsync(n => n.Id == notificacionId &&
                               (!n.FechaExpiracion.HasValue || n.FechaExpiracion > now) &&
                               (n.EsParaTodos || _context.NotificacionDestinatarios.Any(d => d.NotificacionId == n.Id && d.UsuarioId == usuarioId)),
                    cancellationToken);

            var isHidden = await _context.NotificacionOcultas
                .AsNoTracking()
                .AnyAsync(o => o.NotificacionId == notificacionId && o.UsuarioId == usuarioId, cancellationToken);

            if (!hasAccess || isHidden)
            {
                throw new InvalidOperationException("Notificación inexistente o sin acceso");
            }

            var exists = await _context.NotificacionLeidas
                .AnyAsync(x => x.NotificacionId == notificacionId && x.UsuarioId == usuarioId, cancellationToken);

            if (exists)
            {
                return;
            }

            _context.NotificacionLeidas.Add(new NotificacionLeida
            {
                NotificacionId = notificacionId,
                UsuarioId = usuarioId,
                FechaLeida = now
            });

            await _context.SaveChangesAsync(cancellationToken);
        }

        public async Task<NotificacionDetalleDTO> ObtenerDetalleAsync(int notificacionId, CancellationToken cancellationToken = default)
        {
            var usuarioId = GetCurrentUsuarioIdOrThrow();
            var now = DateTime.Now;

            var detalle = await _context.Notificaciones
                .AsNoTracking()
                .Where(n => n.Id == notificacionId)
                .Where(n => !n.FechaExpiracion.HasValue || n.FechaExpiracion > now)
                .Where(n =>
                    n.EsParaTodos ||
                    _context.NotificacionDestinatarios.Any(d => d.NotificacionId == n.Id && d.UsuarioId == usuarioId))
                .Where(n => !_context.NotificacionOcultas.Any(o => o.NotificacionId == n.Id && o.UsuarioId == usuarioId))
                .Select(n => new NotificacionDetalleDTO
                {
                    Id = n.Id,
                    DetalleHtml = n.DetalleHtml
                })
                .FirstOrDefaultAsync(cancellationToken);

            if (detalle == null)
            {
                throw new InvalidOperationException("Notificación inexistente o sin acceso");
            }

            return detalle;
        }

        public async Task OcultarParaUsuarioActualAsync(int notificacionId, CancellationToken cancellationToken = default)
        {
            var usuarioId = GetCurrentUsuarioIdOrThrow();
            var now = DateTime.Now;

            var hasAccess = await _context.Notificaciones
                .AsNoTracking()
                .AnyAsync(n => n.Id == notificacionId &&
                               (!n.FechaExpiracion.HasValue || n.FechaExpiracion > now) &&
                               (n.EsParaTodos || _context.NotificacionDestinatarios.Any(d => d.NotificacionId == n.Id && d.UsuarioId == usuarioId)),
                    cancellationToken);

            if (!hasAccess)
            {
                throw new InvalidOperationException("Notificación inexistente o sin acceso");
            }

            var exists = await _context.NotificacionOcultas
                .AnyAsync(x => x.NotificacionId == notificacionId && x.UsuarioId == usuarioId, cancellationToken);

            if (exists)
            {
                return;
            }

            _context.NotificacionOcultas.Add(new NotificacionOculta
            {
                NotificacionId = notificacionId,
                UsuarioId = usuarioId,
                FechaOculta = now
            });

            await _context.SaveChangesAsync(cancellationToken);
        }

        public async Task<List<NotificacionAdminItemDTO>> AdminListAsync(PaginationDTO pagination, CancellationToken cancellationToken = default)
        {
            GetCurrentUsuarioIdOrThrow();

            var queryable = _context.Notificaciones
                .AsNoTracking()
                .OrderByDescending(x => x.AddRecord)
                .Paginate(pagination);

            var items = await queryable
                .Select(n => new
                {
                    n.Id,
                    n.Titulo,
                    n.Resumen,
                    n.EsParaTodos,
                    n.CreadorUsuarioId,
                    n.AddRecord,
                    n.FechaExpiracion,
                    n.EsGeneradaPorRegla,
                    n.ReglaSqlId,
                    UsuarioIds = _context.NotificacionDestinatarios
                        .Where(d => d.NotificacionId == n.Id)
                        .Select(d => d.UsuarioId)
                        .ToList()
                })
                .ToListAsync(cancellationToken);

            var allUserIds = items
                .SelectMany(x => x.UsuarioIds)
                .Distinct()
                .ToList();

            var userNames = await _context.Usuarios
                .AsNoTracking()
                .Where(u => allUserIds.Contains(u.Id))
                .Select(u => new { u.Id, u.NombreUsuario })
                .ToDictionaryAsync(x => x.Id, x => x.NombreUsuario, cancellationToken);

            return items.Select(x => new NotificacionAdminItemDTO
            {
                Id = x.Id,
                Titulo = x.Titulo,
                Resumen = x.Resumen,
                EsParaTodos = x.EsParaTodos,
                CreadorUsuarioId = x.CreadorUsuarioId,
                AddRecord = x.AddRecord,
                FechaExpiracion = x.FechaExpiracion,
                EsGeneradaPorRegla = x.EsGeneradaPorRegla,
                ReglaSqlId = x.ReglaSqlId,
                UsuarioIds = x.UsuarioIds,
                DestinatariosResumen = x.EsParaTodos
                    ? "Todos"
                    : string.Join(", ", x.UsuarioIds.Select(uid => userNames.TryGetValue(uid, out var name) ? name : uid.ToString()))
            }).ToList();
        }

        public async Task<double> AdminTotalPagesAsync(PaginationDTO pagination, CancellationToken cancellationToken = default)
        {
            GetCurrentUsuarioIdOrThrow();
            double count = await _context.Notificaciones.AsNoTracking().CountAsync(cancellationToken);
            return Math.Ceiling(count / pagination.RecordsNumber);
        }

        public async Task<NotificacionAdminUpsertDTO?> AdminGetByIdAsync(int id, CancellationToken cancellationToken = default)
        {
            GetCurrentUsuarioIdOrThrow();

            var entity = await _context.Notificaciones
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.Id == id, cancellationToken);

            if (entity == null)
            {
                return null;
            }

            var usuarioIds = await _context.NotificacionDestinatarios
                .AsNoTracking()
                .Where(d => d.NotificacionId == id)
                .Select(d => d.UsuarioId)
                .ToListAsync(cancellationToken);

            return new NotificacionAdminUpsertDTO
            {
                Id = entity.Id,
                Titulo = entity.Titulo,
                Resumen = entity.Resumen,
                DetalleHtml = entity.DetalleHtml,
                EsParaTodos = entity.EsParaTodos,
                UsuarioIds = usuarioIds,
                FechaExpiracion = entity.FechaExpiracion
            };
        }

        public async Task<int> AdminCreateAsync(NotificacionAdminUpsertDTO dto, CancellationToken cancellationToken = default)
        {
            var creadorUsuarioId = GetCurrentUsuarioIdOrThrow();

            var usuarioIds = (dto.UsuarioIds ?? new List<int>())
                .Where(x => x > 0)
                .Distinct()
                .ToList();

            if (!dto.EsParaTodos && usuarioIds.Count == 0)
            {
                throw new InvalidOperationException("Debe indicar destinatarios (UsuarioIds) o marcar EsParaTodos");
            }

            var entity = new Notificacion
            {
                Titulo = dto.Titulo,
                Resumen = dto.Resumen,
                DetalleHtml = dto.DetalleHtml,
                FechaExpiracion = dto.FechaExpiracion,
                EsParaTodos = dto.EsParaTodos,
                CreadorUsuarioId = creadorUsuarioId,
                EsGeneradaPorRegla = false,
                ReglaSqlId = null
            };

            _context.Notificaciones.Add(entity);
            await _context.SaveChangesAsync(cancellationToken);

            if (!dto.EsParaTodos)
            {
                _context.NotificacionDestinatarios.AddRange(usuarioIds.Select(uid => new NotificacionDestinatario
                {
                    NotificacionId = entity.Id,
                    UsuarioId = uid
                }));
                await _context.SaveChangesAsync(cancellationToken);
            }

            return entity.Id;
        }

        public async Task AdminUpdateAsync(int id, NotificacionAdminUpsertDTO dto, CancellationToken cancellationToken = default)
        {
            GetCurrentUsuarioIdOrThrow();

            var entity = await _context.Notificaciones
                .FirstOrDefaultAsync(x => x.Id == id, cancellationToken);

            if (entity == null)
            {
                throw new InvalidOperationException("Notificación no encontrada");
            }

            entity.Titulo = dto.Titulo;
            entity.Resumen = dto.Resumen;
            entity.DetalleHtml = dto.DetalleHtml;
            entity.FechaExpiracion = dto.FechaExpiracion;
            entity.EsParaTodos = dto.EsParaTodos;

            await _context.SaveChangesAsync(cancellationToken);

            var newUsuarioIds = (dto.UsuarioIds ?? new List<int>())
                .Where(x => x > 0)
                .Distinct()
                .ToList();

            if (!dto.EsParaTodos && newUsuarioIds.Count == 0)
            {
                throw new InvalidOperationException("Debe indicar destinatarios (UsuarioIds) o marcar EsParaTodos");
            }

            var currentDest = await _context.NotificacionDestinatarios
                .Where(d => d.NotificacionId == id)
                .ToListAsync(cancellationToken);

            _context.NotificacionDestinatarios.RemoveRange(currentDest);
            await _context.SaveChangesAsync(cancellationToken);

            if (!dto.EsParaTodos)
            {
                _context.NotificacionDestinatarios.AddRange(newUsuarioIds.Select(uid => new NotificacionDestinatario
                {
                    NotificacionId = id,
                    UsuarioId = uid
                }));
                await _context.SaveChangesAsync(cancellationToken);
            }
        }

        public async Task AdminDeleteAsync(int id, CancellationToken cancellationToken = default)
        {
            try
            {
                GetCurrentUsuarioIdOrThrow();

                var entity = await _context.Notificaciones
                    .FirstOrDefaultAsync(x => x.Id == id, cancellationToken);

                if (entity == null)
                {
                    return;
                }

                var dest = await _context.NotificacionDestinatarios
                    .Where(d => d.NotificacionId == id)
                    .ToListAsync(cancellationToken);

                _context.NotificacionLeidas.RemoveRange(_context.NotificacionLeidas.Where(l => l.NotificacionId == id));
                _context.NotificacionOcultas.RemoveRange(_context.NotificacionOcultas.Where(o => o.NotificacionId == id));
                _context.NotificacionDestinatarios.RemoveRange(dest);
                _context.Notificaciones.Remove(entity);

                await _context.SaveChangesAsync(cancellationToken);
            }
            catch (Exception ex)
            {
                // Aquí puedes registrar el error (Log)
                // _logger.LogError(ex, "Error al eliminar la notificación {Id}", id);

                throw; // Vuelve a lanzar la excepción para que la maneje el controlador o el middleware global.
            }
        }
    }
}
