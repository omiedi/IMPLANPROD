using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/seguridad-rutas")]
    [Authorize]
    public class SeguridadRutasController : ControllerBase
    {
        private readonly ISeguridadRutasService _seguridadRutasService;
        private readonly IUserContextService _userContextService;
        private readonly DataContext _context;

        public SeguridadRutasController(
            ISeguridadRutasService seguridadRutasService,
            IUserContextService userContextService,
            DataContext context)
        {
            _seguridadRutasService = seguridadRutasService;
            _userContextService = userContextService;
            _context = context;
        }

        [HttpGet("bootstrap-superadmin")]
        [AllowAnonymous]
        public async Task<ActionResult<BootstrapSuperAdminStatusDto>> GetBootstrapSuperAdminStatus(CancellationToken cancellationToken)
        {
            var totalSuperAdmins = await _seguridadRutasService.CountSuperAdminsAsync(cancellationToken);
            return Ok(new BootstrapSuperAdminStatusDto
            {
                TotalSuperAdmins = totalSuperAdmins
            });
        }

        [HttpGet("modulos")]
        public async Task<ActionResult<List<SeguridadModuloDto>>> GetModulos([FromQuery] bool incluirConfiguracion = false, CancellationToken cancellationToken = default)
        {
            var modulos = await _seguridadRutasService.GetModulosAsync(incluirConfiguracion, cancellationToken);
            return Ok(modulos);
        }

        [HttpGet("modulos/{moduloId:int}/paginas")]
        public async Task<ActionResult<List<SeguridadPaginaDto>>> GetPaginasByModulo(int moduloId, [FromQuery] bool incluirConfiguracion = false, CancellationToken cancellationToken = default)
        {
            var paginas = await _seguridadRutasService.GetPaginasByModuloAsync(moduloId, incluirConfiguracion, cancellationToken);
            return Ok(paginas);
        }

        [HttpGet("usuarios/{usuarioId:int}/paginas")]
        public async Task<ActionResult<List<int>>> GetPaginasPermitidasByUsuario(int usuarioId, CancellationToken cancellationToken = default)
        {
            var paginaIds = await _seguridadRutasService.GetPaginasPermitidasByUsuarioAsync(usuarioId, cancellationToken);
            return Ok(paginaIds);
        }

        [HttpPost("usuarios/{usuarioId:int}/paginas")]
        public async Task<ActionResult> AsignarPaginasUsuario(int usuarioId, [FromBody] AsignarPaginasUsuarioRequestDto request, CancellationToken cancellationToken = default)
        {
            var isCurrentSuperAdmin = User.IsInRole("SUPERADMIN");
            var isCurrentAdmin      = User.IsInRole("ADMINISTRADOR") || User.IsInRole("ADMIN");

            if (!isCurrentSuperAdmin && !isCurrentAdmin)
            {
                return Forbid();
            }

            var targetUser = await _context.Usuarios
                .AsNoTracking()
                .Include(x => x.Perfil)
                .FirstOrDefaultAsync(x => x.Id == usuarioId, cancellationToken);

            if (targetUser == null)
            {
                return NotFound("Usuario objetivo no encontrado.");
            }

            var targetPerfil = (targetUser.Perfil?.Nombre ?? string.Empty).Trim();
            var isTargetSuperAdmin = string.Equals(targetPerfil, "SUPERADMIN", StringComparison.OrdinalIgnoreCase);

            if (isTargetSuperAdmin && !isCurrentSuperAdmin)
            {
                return BadRequest("Solo un SUPERADMIN puede modificar permisos de páginas de otro SUPERADMIN.");
            }

            await _seguridadRutasService.AsignarPaginasUsuarioAsync(usuarioId, request.PaginaIds, cancellationToken);
            return Ok();
        }

        [HttpPost("usuarios/{origenId:int}/copiar-permisos/{destinoId:int}")]
        public async Task<ActionResult> CopiarPermisosUsuario(int origenId, int destinoId, CancellationToken cancellationToken = default)
        {
            var isCurrentSuperAdmin = User.IsInRole("SUPERADMIN");
            var isCurrentAdmin      = User.IsInRole("ADMINISTRADOR") || User.IsInRole("ADMIN");

            if (!isCurrentSuperAdmin && !isCurrentAdmin)
                return Forbid();

            if (origenId == destinoId)
                return BadRequest("El usuario origen y destino no pueden ser el mismo.");

            var origenExiste  = await _context.Usuarios.AsNoTracking().AnyAsync(x => x.Id == origenId,  cancellationToken);
            var destinoExiste = await _context.Usuarios.AsNoTracking().AnyAsync(x => x.Id == destinoId, cancellationToken);

            if (!origenExiste)  return NotFound($"Usuario origen (Id={origenId}) no encontrado.");
            if (!destinoExiste) return NotFound($"Usuario destino (Id={destinoId}) no encontrado.");

            var paginaIds = await _seguridadRutasService.GetPaginasPermitidasByUsuarioAsync(origenId, cancellationToken);
            await _seguridadRutasService.AsignarPaginasUsuarioAsync(destinoId, paginaIds, cancellationToken);

            return Ok();
        }

        [HttpGet("usuarios/{usuarioId:int}/puede-acceder")]
        public async Task<ActionResult<object>> PuedeAccederRuta(int usuarioId, [FromQuery] string ruta, CancellationToken cancellationToken = default)
        {
            var puedeAcceder = await _seguridadRutasService.PuedeAccederRutaAsync(usuarioId, ruta, cancellationToken);
            return Ok(new
            {
                usuarioId,
                ruta,
                puedeAcceder
            });
        }
    }
}
