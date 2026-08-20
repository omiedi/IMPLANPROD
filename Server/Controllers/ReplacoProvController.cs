using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/replacoprov")]
    public class ReplacoProvController : ControllerBase
    {
        private readonly DataContext _context;

        public ReplacoProvController(DataContext context)
        {
            _context = context;
        }

        [HttpGet("proveedor/{numeProv:int}")]
        public async Task<ActionResult<List<ReplacoProvRowDTO>>> GetByProveedor(int numeProv)
        {
            var query =
                from r in _context.Replacos.AsNoTracking()
                where r.NumeProv == numeProv
                join m in _context.masters.AsNoTracking() on r.CodInte equals m.Codint into mj
                from m in mj.DefaultIfEmpty()
                orderby (m.Descripcion ?? "")
                select new ReplacoProvRowDTO
                {
                    Id = r.Id,
                    NumeProv = r.NumeProv,
                    CodInte = r.CodInte,
                    Codigo = m.Codigo,
                    Descripcion = m.Descripcion,
                    Referencia = r.Referencia,
                    Observaciones = r.Observaciones
                };

            return Ok(await query.ToListAsync());
        }

        [HttpGet("{id:int}")]
        public async Task<ActionResult<Replaco>> GetById(int id)
        {
            var entity = await _context.Replacos.FirstOrDefaultAsync(x => x.Id == id);
            if (entity == null)
            {
                return NotFound();
            }

            return Ok(entity);
        }

        [HttpPost]
        public async Task<ActionResult> PostAsync(Replaco model)
        {
            if (model.NumeProv == null || model.NumeProv <= 0)
            {
                return BadRequest("Debe seleccionar un proveedor.");
            }

            if (model.CodInte <= 0)
            {
                return BadRequest("Debe seleccionar un código de Master.");
            }

            model.NumeCli = 0;

            var exists = await _context.Replacos.AnyAsync(x => x.NumeProv == model.NumeProv && x.CodInte == model.CodInte);
            if (exists)
            {
                return BadRequest("El código de Master ya existe para este proveedor.");
            }

            _context.Replacos.Add(model);
            await _context.SaveChangesAsync();

            return Ok(model);
        }

        [HttpPut]
        public async Task<ActionResult> PutAsync(Replaco model)
        {
            if (model.Id <= 0)
            {
                return BadRequest("Id inválido.");
            }

            if (model.NumeProv == null || model.NumeProv <= 0)
            {
                return BadRequest("Debe seleccionar un proveedor.");
            }

            if (model.CodInte <= 0)
            {
                return BadRequest("Debe seleccionar un código de Master.");
            }

            model.NumeCli = 0;

            var exists = await _context.Replacos.AnyAsync(x => x.Id != model.Id && x.NumeProv == model.NumeProv && x.CodInte == model.CodInte);
            if (exists)
            {
                return BadRequest("El código de Master ya existe para este proveedor.");
            }

            _context.Replacos.Update(model);
            await _context.SaveChangesAsync();

            return Ok(model);
        }

        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var entity = await _context.Replacos.FirstOrDefaultAsync(x => x.Id == id);
            if (entity == null)
            {
                return NotFound();
            }

            _context.Replacos.Remove(entity);
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
