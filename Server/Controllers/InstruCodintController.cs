using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/instru-codint")]
    public class InstruCodintController : ControllerBase
    {
        private readonly DataContext _context;

        public InstruCodintController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene las instrucciones guardadas para un producto
        /// GET /api/instru-codint/by-master/{masterCodInte}
        /// </summary>
        [HttpGet("by-master/{masterCodInte:int}")]
        public async Task<ActionResult<List<Instru_Codint>>> GetByMaster(int masterCodInte)
        {
            var lista = await _context.Instru_Codints
                .AsNoTracking()
                .Where(x => x.Master_Cod_Inte == masterCodInte && x.Status != -1)
                .ToListAsync();
            return lista;
        }

        /// <summary>
        /// Guarda (reemplaza) las instrucciones para un producto
        /// POST /api/instru-codint/guardar/{masterCodInte}
        /// </summary>
        [HttpPost("guardar/{masterCodInte:int}")]
        public async Task<ActionResult> Guardar(int masterCodInte, [FromBody] List<Instru_Codint> items)
        {
            // Eliminar físicamente los registros anteriores de este producto
            var existentes = await _context.Instru_Codints
                .Where(x => x.Master_Cod_Inte == masterCodInte)
                .ToListAsync();

            if (existentes.Any())
                _context.Instru_Codints.RemoveRange(existentes);

            // Insertar los nuevos
            foreach (var item in items)
            {
                item.Master_Cod_Inte = masterCodInte;
                item.Status          = 0;
                item.AddRecord       = DateTime.Now;
                item.LastUpdate      = DateTime.Now;
                _context.Instru_Codints.Add(item);
            }

            await _context.SaveChangesAsync();
            return Ok(new { guardados = items.Count });
        }
    }
}
