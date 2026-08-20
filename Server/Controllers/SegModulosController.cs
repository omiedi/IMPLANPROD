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
    public class SegModulosController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IUserContextService _userContextService;

        public SegModulosController(DataContext context, IUserContextService userContextService)
        {
            _context = context;
            _userContextService = userContextService;
        }

        [HttpGet]
        public async Task<ActionResult<List<SegModulo>>> GetAll(CancellationToken cancellationToken = default)
        {
            return await _context.SegModulos
                .AsNoTracking()
                .OrderBy(x => x.Orden)
                .ThenBy(x => x.Nombre)
                .ToListAsync(cancellationToken);
        }

        [HttpGet("{id:int}")]
        public async Task<ActionResult<SegModulo>> GetById(int id, CancellationToken cancellationToken = default)
        {
            var entity = await _context.SegModulos
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.Id == id, cancellationToken);

            if (entity == null) return NotFound();
            return Ok(entity);
        }

        [HttpPost]
        public async Task<ActionResult<SegModulo>> Create([FromBody] SegModulo model, CancellationToken cancellationToken = default)
        {
            if (await _context.SegModulos.AnyAsync(x => x.Codigo == model.Codigo, cancellationToken))
                return BadRequest($"Ya existe un módulo con el código '{model.Codigo}'.");

            model.Id = 0;
            model.Codiempr = 0;
            model.CodiEmprNet = 0;
            model.NumUsuar = _userContextService.GetCurrentUsuarioId();
            model.AddRecord = DateTime.UtcNow;
            model.LastUpdate = DateTime.UtcNow;

            _context.SegModulos.Add(model);
            await _context.SaveChangesAsync(cancellationToken);
            return CreatedAtAction(nameof(GetById), new { id = model.Id }, model);
        }

        [HttpPut("{id:int}")]
        public async Task<ActionResult<SegModulo>> Update(int id, [FromBody] SegModulo model, CancellationToken cancellationToken = default)
        {
            if (id != model.Id) return BadRequest("El Id no coincide.");

            var entity = await _context.SegModulos.FirstOrDefaultAsync(x => x.Id == id, cancellationToken);
            if (entity == null) return NotFound();

            if (await _context.SegModulos.AnyAsync(x => x.Codigo == model.Codigo && x.Id != id, cancellationToken))
                return BadRequest($"Ya existe otro módulo con el código '{model.Codigo}'.");

            entity.Codigo = model.Codigo;
            entity.Nombre = model.Nombre;
            entity.Descripcion = model.Descripcion;
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
            var entity = await _context.SegModulos.FirstOrDefaultAsync(x => x.Id == id, cancellationToken);
            if (entity == null) return NotFound();

            var tienePaginas = await _context.SegPaginas.AnyAsync(x => x.ModuloId == id, cancellationToken);
            if (tienePaginas)
                return BadRequest("No se puede eliminar el módulo porque tiene páginas asociadas. Elimine primero las páginas.");

            _context.SegModulos.Remove(entity);
            await _context.SaveChangesAsync(cancellationToken);
            return NoContent();
        }
    }
}
