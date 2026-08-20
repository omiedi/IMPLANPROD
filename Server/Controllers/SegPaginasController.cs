using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class SegPaginasController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IUserContextService _userContextService;

        public SegPaginasController(DataContext context, IUserContextService userContextService)
        {
            _context = context;
            _userContextService = userContextService;
        }

        [HttpGet]
        public async Task<ActionResult<List<SegPagina>>> GetAll([FromQuery] int? moduloId = null, CancellationToken cancellationToken = default)
        {
            var query = _context.SegPaginas
                .AsNoTracking()
                .Include(x => x.Modulo)
                .AsQueryable();

            if (moduloId.HasValue && moduloId.Value > 0)
                query = query.Where(x => x.ModuloId == moduloId.Value);

            return await query
                .OrderBy(x => x.ModuloId)
                .ThenBy(x => x.Orden)
                .ThenBy(x => x.Nombre)
                .ToListAsync(cancellationToken);
        }

        [HttpGet("{id:int}")]
        public async Task<ActionResult<SegPagina>> GetById(int id, CancellationToken cancellationToken = default)
        {
            var entity = await _context.SegPaginas
                .AsNoTracking()
                .Include(x => x.Modulo)
                .FirstOrDefaultAsync(x => x.Id == id, cancellationToken);

            if (entity == null) return NotFound();
            return Ok(entity);
        }

        [HttpGet("modulos-combo")]
        public async Task<ActionResult<List<SegModulo>>> GetModulosCombo(CancellationToken cancellationToken = default)
        {
            return await _context.SegModulos
                .AsNoTracking()
                .Where(x => x.Activo)
                .OrderBy(x => x.Orden)
                .ThenBy(x => x.Nombre)
                .ToListAsync(cancellationToken);
        }

        [HttpPost]
        public async Task<ActionResult<SegPagina>> Create([FromBody] SegPagina model, CancellationToken cancellationToken = default)
        {
            if (await _context.SegPaginas.AnyAsync(x => x.Ruta == model.Ruta, cancellationToken))
                return BadRequest($"Ya existe una página con la ruta '{model.Ruta}'.");

            model.Id = 0;
            model.Codiempr = 0;
            model.CodiEmprNet = 0;
            model.NumUsuar = _userContextService.GetCurrentUsuarioId();
            model.AddRecord = DateTime.UtcNow;
            model.LastUpdate = DateTime.UtcNow;
            model.Modulo = null;

            _context.SegPaginas.Add(model);
            await _context.SaveChangesAsync(cancellationToken);
            return CreatedAtAction(nameof(GetById), new { id = model.Id }, model);
        }

        [HttpPut("{id:int}")]
        public async Task<ActionResult<SegPagina>> Update(int id, [FromBody] SegPagina model, CancellationToken cancellationToken = default)
        {
            if (id != model.Id) return BadRequest("El Id no coincide.");

            var entity = await _context.SegPaginas.FirstOrDefaultAsync(x => x.Id == id, cancellationToken);
            if (entity == null) return NotFound();

            if (await _context.SegPaginas.AnyAsync(x => x.Ruta == model.Ruta && x.Id != id, cancellationToken))
                return BadRequest($"Ya existe otra página con la ruta '{model.Ruta}'.");

            entity.ModuloId = model.ModuloId;
            entity.Nombre = model.Nombre;
            entity.Ruta = model.Ruta;
            entity.Componente = model.Componente;
            entity.EsConfiguracion = model.EsConfiguracion;
            entity.Orden = model.Orden;
            entity.Activo = model.Activo;
            entity.LastUpdate = DateTime.UtcNow;
            entity.NumUsuar = _userContextService.GetCurrentUsuarioId();

            await _context.SaveChangesAsync(cancellationToken);
            return Ok(entity);
        }

        [HttpDelete("{id:int}")]
        public async Task<IActionResult> Delete(int id, CancellationToken cancellationToken = default)
        {
            var entity = await _context.SegPaginas.FirstOrDefaultAsync(x => x.Id == id, cancellationToken);
            if (entity == null) return NotFound();

            var tieneAsignaciones = await _context.SegUsuarioPaginas.AnyAsync(x => x.PaginaId == id, cancellationToken);
            if (tieneAsignaciones)
                return BadRequest("No se puede eliminar la página porque tiene asignaciones de usuarios. Quite primero los permisos.");

            _context.SegPaginas.Remove(entity);
            await _context.SaveChangesAsync(cancellationToken);
            return NoContent();
        }
    }
}
