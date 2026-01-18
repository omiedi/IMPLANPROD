using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using System.Collections.Generic;
using System.Threading.Tasks;
using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PerfilesController : ControllerBase
    {
        private readonly DataContext _context;

        public PerfilesController(DataContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Perfil>>> GetPerfiles()
        {
            return await _context.Perfiles.Include(p => p.Permisos).ToListAsync();
        }

        [HttpPost]
        public async Task<ActionResult<Perfil>> CrearPerfil(Perfil perfil)
        {
            _context.Perfiles.Add(perfil);
            await _context.SaveChangesAsync();

            return CreatedAtAction(nameof(GetPerfiles), new { id = perfil.Id }, perfil);
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> ActualizarPerfil(int id, Perfil perfil)
        {
            if (id != perfil.Id)
            {
                return BadRequest();
            }

            _context.Entry(perfil).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!PerfilExists(id))
                {
                    return NotFound();
                }
                throw;
            }

            return NoContent();
        }

        private bool PerfilExists(int id)
        {
            return _context.Perfiles.Any(e => e.Id == id);
        }
    }
}
