using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Linq;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controller para el modulo Seguimiento de Solicitudes
    /// </summary>
    [ApiController]
    [Route("/api/solicitudes")]
    [Authorize]
    public class SolicitudesController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IUserContextService _userContextService;

        public SolicitudesController(DataContext context, IUserContextService userContextService)
        {
            _context = context;
            _userContextService = userContextService;
        }

        private bool EsSuperAdmin() => User.IsInRole("SUPERADMIN");
        private bool EsAdministrador() => User.IsInRole("SUPERADMIN") || User.IsInRole("ADMINISTRADOR") || User.IsInRole("ADMIN");

        /// <summary>
        /// Listado de solicitudes del usuario actual (solo las suyas)
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<List<SolicitudListadoDTO>>> Get([FromQuery] SolicitudFiltroDTO filtro)
        {
            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return Unauthorized();

            var query = _context.Solicitudes
                .AsNoTracking()
                .Where(x => x.UsuarioSolicitante == userId && x.CodiEmprNet == codiEmpr)
                .AsQueryable();

            query = AplicarFiltros(query, filtro);

            var usuarios = await ObtenerNombresUsuariosAsync(query.Select(x => x.UsuarioSolicitante).Concat(query.Select(x => x.UsuarioUltimaActualizacion ?? 0)));

            var result = await query
                .OrderByDescending(x => x.FechaIngreso)
                .Paginate(filtro)
                .Select(x => new SolicitudListadoDTO
                {
                    Id = x.Id,
                    FechaIngreso = x.FechaIngreso,
                    Titulo = x.Titulo,
                    Modulo = x.Modulo,
                    TipoSolicitud = x.TipoSolicitud,
                    Nivel = x.Nivel,
                    Estado = x.Estado,
                    FechaUltimaActualizacion = x.FechaUltimaActualizacion,
                    UsuarioUltimaActualizacion = x.UsuarioUltimaActualizacion,
                    UsuarioSolicitante = x.UsuarioSolicitante,
                    NombreSolicitante = usuarios.ContainsKey(x.UsuarioSolicitante) ? usuarios[x.UsuarioSolicitante] : $"Usuario {x.UsuarioSolicitante}"
                })
                .ToListAsync();

            return Ok(result);
        }

        /// <summary>
        /// Total de paginas para el listado del usuario actual
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult<double>> GetTotalPages([FromQuery] SolicitudFiltroDTO filtro)
        {
            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return Unauthorized();

            var query = _context.Solicitudes
                .AsNoTracking()
                .Where(x => x.UsuarioSolicitante == userId && x.CodiEmprNet == codiEmpr)
                .AsQueryable();

            query = AplicarFiltros(query, filtro);

            double count = await query.CountAsync();
            double totalPages = Math.Ceiling(count / filtro.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Listado administrativo de solicitudes. Permite ver todas las de la empresa
        /// o todas las del sistema si es SUPERADMIN y activa VerTodos.
        /// </summary>
        [HttpGet("admin")]
        public async Task<ActionResult<List<SolicitudListadoDTO>>> GetAdmin([FromQuery] SolicitudFiltroDTO filtro)
        {
            if (!EsAdministrador())
                return Forbid();

            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return Unauthorized();

            var query = _context.Solicitudes.AsNoTracking().AsQueryable();

            if (!filtro.VerTodos || !EsSuperAdmin())
                query = query.Where(x => x.CodiEmprNet == codiEmpr);

            if (filtro.UsuarioFiltro.HasValue)
                query = query.Where(x => x.UsuarioSolicitante == filtro.UsuarioFiltro.Value);

            query = AplicarFiltros(query, filtro);

            var idsSolicitantes = await query.Select(x => x.UsuarioSolicitante).ToListAsync();
            var idsUltimos = await query.Select(x => x.UsuarioUltimaActualizacion ?? 0).ToListAsync();
            var usuarios = await ObtenerNombresUsuariosAsync(idsSolicitantes.Concat(idsUltimos));

            var result = await query
                .OrderByDescending(x => x.FechaIngreso)
                .Paginate(filtro)
                .Select(x => new SolicitudListadoDTO
                {
                    Id = x.Id,
                    FechaIngreso = x.FechaIngreso,
                    Titulo = x.Titulo,
                    Modulo = x.Modulo,
                    TipoSolicitud = x.TipoSolicitud,
                    Nivel = x.Nivel,
                    Estado = x.Estado,
                    FechaUltimaActualizacion = x.FechaUltimaActualizacion,
                    UsuarioUltimaActualizacion = x.UsuarioUltimaActualizacion,
                    UsuarioSolicitante = x.UsuarioSolicitante,
                    NombreSolicitante = usuarios.ContainsKey(x.UsuarioSolicitante) ? usuarios[x.UsuarioSolicitante] : $"Usuario {x.UsuarioSolicitante}"
                })
                .ToListAsync();

            return Ok(result);
        }

        /// <summary>
        /// Total de paginas para el listado administrativo
        /// </summary>
        [HttpGet("admin/totalPages")]
        public async Task<ActionResult<double>> GetAdminTotalPages([FromQuery] SolicitudFiltroDTO filtro)
        {
            if (!EsAdministrador())
                return Forbid();

            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return Unauthorized();

            var query = _context.Solicitudes.AsNoTracking().AsQueryable();

            if (!filtro.VerTodos || !EsSuperAdmin())
                query = query.Where(x => x.CodiEmprNet == codiEmpr);

            if (filtro.UsuarioFiltro.HasValue)
                query = query.Where(x => x.UsuarioSolicitante == filtro.UsuarioFiltro.Value);

            query = AplicarFiltros(query, filtro);

            double count = await query.CountAsync();
            double totalPages = Math.Ceiling(count / filtro.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene el detalle de una solicitud con su historial de avances
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<ActionResult<SolicitudDetalleDTO>> GetById(int id)
        {
            var solicitud = await _context.Solicitudes
                .AsNoTracking()
                .Include(x => x.Avances)
                .FirstOrDefaultAsync(x => x.Id == id);

            if (solicitud == null)
                return NotFound();

            if (!PuedeConsultar(solicitud))
                return Forbid();

            var usuarios = await ObtenerNombresUsuariosAsync(new[] { solicitud.UsuarioSolicitante }
                .Concat(solicitud.Avances.Select(a => a.Usuario))
                .Concat(solicitud.UsuarioUltimaActualizacion.HasValue ? new[] { solicitud.UsuarioUltimaActualizacion.Value } : Array.Empty<int>())
                .Concat(solicitud.UsuarioCierre.HasValue ? new[] { solicitud.UsuarioCierre.Value } : Array.Empty<int>()));

            return Ok(new SolicitudDetalleDTO
            {
                Id = solicitud.Id,
                FechaIngreso = solicitud.FechaIngreso,
                UsuarioSolicitante = solicitud.UsuarioSolicitante,
                NombreSolicitante = usuarios.ContainsKey(solicitud.UsuarioSolicitante) ? usuarios[solicitud.UsuarioSolicitante] : $"Usuario {solicitud.UsuarioSolicitante}",
                Modulo = solicitud.Modulo,
                TipoSolicitud = solicitud.TipoSolicitud,
                Nivel = solicitud.Nivel,
                Titulo = solicitud.Titulo,
                Detalle = solicitud.Detalle,
                Estado = solicitud.Estado,
                FechaUltimaActualizacion = solicitud.FechaUltimaActualizacion,
                UsuarioUltimaActualizacion = solicitud.UsuarioUltimaActualizacion,
                FechaCierre = solicitud.FechaCierre,
                UsuarioCierre = solicitud.UsuarioCierre,
                Avances = solicitud.Avances
                    .OrderBy(a => a.Fecha)
                    .Select(a => new SolicitudAvanceDTO
                    {
                        Id = a.Id,
                        SolicitudId = a.SolicitudId,
                        Fecha = a.Fecha,
                        Estado = a.Estado,
                        Comentario = a.Comentario,
                        Usuario = a.Usuario,
                        NombreUsuario = usuarios.ContainsKey(a.Usuario) ? usuarios[a.Usuario] : $"Usuario {a.Usuario}"
                    })
                    .ToList()
            });
        }

        /// <summary>
        /// Crea una nueva solicitud con estado Ingresado y un avance inicial
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<SolicitudDetalleDTO>> Post(CrearSolicitudDTO dto)
        {
            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return Unauthorized();

            var ahora = DateTime.Now;

            int solicitante = (EsAdministrador() && dto.UsuarioSolicitante.HasValue && dto.UsuarioSolicitante.Value > 0)
                ? dto.UsuarioSolicitante.Value
                : userId.Value;

            var solicitud = new Solicitud
            {
                FechaIngreso = ahora,
                UsuarioSolicitante = solicitante,
                Modulo = dto.Modulo,
                TipoSolicitud = dto.TipoSolicitud,
                Nivel = dto.Nivel,
                Titulo = dto.Titulo,
                Detalle = dto.Detalle,
                Estado = "Ingresado",
                FechaUltimaActualizacion = ahora,
                UsuarioUltimaActualizacion = solicitante,
                CodiEmprNet = codiEmpr.Value,
                NumUsuar = userId.Value,
                Avances = new List<SolicitudAvance>
                {
                    new SolicitudAvance
                    {
                        Fecha = ahora,
                        Estado = "Ingresado",
                        Comentario = "Solicitud registrada.",
                        Usuario = solicitante,
                        CodiEmprNet = codiEmpr.Value,
                        NumUsuar = userId.Value
                    }
                }
            };

            _context.Solicitudes.Add(solicitud);
            await _context.SaveChangesAsync();

            return await GetById(solicitud.Id);
        }

        /// <summary>
        /// Agrega un avance a una solicitud, actualiza el estado actual y
        /// registra fecha/usuario de cierre si corresponde.
        /// </summary>
        [HttpPost("avance")]
        public async Task<ActionResult<SolicitudDetalleDTO>> PostAvance(CrearSolicitudAvanceDTO dto)
        {
            if (!EsAdministrador())
                return Forbid();

            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return Unauthorized();

            var solicitud = await _context.Solicitudes
                .FirstOrDefaultAsync(x => x.Id == dto.SolicitudId);

            if (solicitud == null)
                return NotFound();

            if (!PuedeModificar(solicitud))
                return Forbid();

            if (solicitud.Estado == "Cerrado" || solicitud.Estado == "Cancelado")
                return BadRequest("No se pueden agregar avances a una solicitud Cerrada o Cancelada.");

            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                var ahora = DateTime.Now;

                var avance = new SolicitudAvance
                {
                    SolicitudId = solicitud.Id,
                    Fecha = ahora,
                    Estado = dto.Estado,
                    Comentario = dto.Comentario,
                    Usuario = userId.Value,
                    CodiEmprNet = codiEmpr.Value,
                    NumUsuar = userId.Value
                };

                _context.SolicitudesAvances.Add(avance);

                solicitud.Estado = dto.Estado;
                solicitud.FechaUltimaActualizacion = ahora;
                solicitud.UsuarioUltimaActualizacion = userId.Value;

                if (dto.Estado == "Cerrado")
                {
                    solicitud.FechaCierre = ahora;
                    solicitud.UsuarioCierre = userId.Value;
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();
            }
            catch
            {
                await transaction.RollbackAsync();
                throw;
            }

            return await GetById(solicitud.Id);
        }

        /// <summary>
        /// Cancela una solicitud como avance de estado
        /// </summary>
        [HttpPost("{id:int}/cancelar")]
        public async Task<ActionResult<SolicitudDetalleDTO>> Cancelar(int id)
        {
            if (!EsAdministrador())
                return Forbid();

            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return Unauthorized();

            var solicitud = await _context.Solicitudes
                .FirstOrDefaultAsync(x => x.Id == id);

            if (solicitud == null)
                return NotFound();

            if (!PuedeModificar(solicitud))
                return Forbid();

            if (solicitud.Estado == "Cerrado" || solicitud.Estado == "Cancelado")
                return BadRequest("La solicitud ya se encuentra Cerrada o Cancelada.");

            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                var ahora = DateTime.Now;

                var avance = new SolicitudAvance
                {
                    SolicitudId = solicitud.Id,
                    Fecha = ahora,
                    Estado = "Cancelado",
                    Comentario = "Solicitud cancelada.",
                    Usuario = userId.Value,
                    CodiEmprNet = codiEmpr.Value,
                    NumUsuar = userId.Value
                };

                _context.SolicitudesAvances.Add(avance);

                solicitud.Estado = "Cancelado";
                solicitud.FechaUltimaActualizacion = ahora;
                solicitud.UsuarioUltimaActualizacion = userId.Value;

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();
            }
            catch
            {
                await transaction.RollbackAsync();
                throw;
            }

            return await GetById(solicitud.Id);
        }

        /// <summary>
        /// Cerrar una solicitud directamente. Equivale a un avance con estado Cerrado.
        /// </summary>
        [HttpPost("{id:int}/cerrar")]
        public async Task<ActionResult<SolicitudDetalleDTO>> Cerrar(int id, [FromBody] string comentario)
        {
            if (!EsAdministrador())
                return Forbid();

            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return Unauthorized();

            var solicitud = await _context.Solicitudes
                .FirstOrDefaultAsync(x => x.Id == id);

            if (solicitud == null)
                return NotFound();

            if (!PuedeModificar(solicitud))
                return Forbid();

            if (solicitud.Estado == "Cerrado" || solicitud.Estado == "Cancelado")
                return BadRequest("La solicitud ya se encuentra Cerrada o Cancelada.");

            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                var ahora = DateTime.Now;

                var avance = new SolicitudAvance
                {
                    SolicitudId = solicitud.Id,
                    Fecha = ahora,
                    Estado = "Cerrado",
                    Comentario = string.IsNullOrWhiteSpace(comentario) ? "Solicitud cerrada." : comentario,
                    Usuario = userId.Value,
                    CodiEmprNet = codiEmpr.Value,
                    NumUsuar = userId.Value
                };

                _context.SolicitudesAvances.Add(avance);

                solicitud.Estado = "Cerrado";
                solicitud.FechaUltimaActualizacion = ahora;
                solicitud.UsuarioUltimaActualizacion = userId.Value;
                solicitud.FechaCierre = ahora;
                solicitud.UsuarioCierre = userId.Value;

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();
            }
            catch
            {
                await transaction.RollbackAsync();
                throw;
            }

            return await GetById(solicitud.Id);
        }

        #region Helpers privados

        private IQueryable<Solicitud> AplicarFiltros(IQueryable<Solicitud> query, SolicitudFiltroDTO filtro)
        {
            if (filtro.FechaDesde.HasValue)
                query = query.Where(x => x.FechaIngreso >= filtro.FechaDesde.Value);

            if (filtro.FechaHasta.HasValue)
                query = query.Where(x => x.FechaIngreso < filtro.FechaHasta.Value.AddDays(1));

            if (!string.IsNullOrWhiteSpace(filtro.Modulo))
                query = query.Where(x => x.Modulo == filtro.Modulo);

            if (!string.IsNullOrWhiteSpace(filtro.TipoSolicitud))
                query = query.Where(x => x.TipoSolicitud == filtro.TipoSolicitud);

            if (!string.IsNullOrWhiteSpace(filtro.Nivel))
                query = query.Where(x => x.Nivel == filtro.Nivel);

            if (!string.IsNullOrWhiteSpace(filtro.Estado))
                query = query.Where(x => x.Estado == filtro.Estado);

            if (filtro.Numero.HasValue)
                query = query.Where(x => x.Id == filtro.Numero.Value);

            return query;
        }

        private bool PuedeConsultar(Solicitud solicitud)
        {
            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return false;

            if (EsSuperAdmin())
                return true;

            if (solicitud.CodiEmprNet != codiEmpr)
                return false;

            if (solicitud.UsuarioSolicitante == userId)
                return true;

            return EsAdministrador();
        }

        private bool PuedeModificar(Solicitud solicitud)
        {
            var userId = _userContextService.GetCurrentUserId();
            var codiEmpr = _userContextService.GetCurrentCodigoEmpresa();
            if (userId == null || codiEmpr == null)
                return false;

            if (EsSuperAdmin())
                return true;

            if (solicitud.CodiEmprNet != codiEmpr)
                return false;

            return EsAdministrador();
        }

        private async Task<Dictionary<int, string>> ObtenerNombresUsuariosAsync(IEnumerable<int> ids)
        {
            var idsUnicos = ids.Distinct().Where(id => id > 0).ToList();
            if (!idsUnicos.Any())
                return new Dictionary<int, string>();

            return await _context.Usuarios
                .AsNoTracking()
                .Where(u => idsUnicos.Contains(u.IdFlexoft))
                .ToDictionaryAsync(
                    u => u.IdFlexoft,
                    u => u.NombreUsuario ?? $"Usuario {u.IdFlexoft}");
        }

        #endregion
    }
}
