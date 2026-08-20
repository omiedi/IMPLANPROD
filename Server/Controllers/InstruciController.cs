using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/instruci")]
    public class InstruciController : ControllerBase
    {
        private readonly DataContext _context;

        public InstruciController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Búsqueda de instrucciones por Codigo o Descripcion.
        /// Si filter está vacío, devuelve todos los registros activos hasta pageSize.
        /// GET /api/instruci/search?filter=texto&pageSize=100
        /// </summary>
        [HttpGet("search")]
        public async Task<ActionResult<List<Instruci>>> Search([FromQuery] string? filter = null, [FromQuery] int pageSize = 100)
        {
            if (pageSize > 200) pageSize = 200;

            var query = _context.Instrucis.AsNoTracking();

            if (!string.IsNullOrWhiteSpace(filter))
            {
                var term = filter.Trim().ToLower();
                query = query.Where(x => 
                    (x.Codigo != null && x.Codigo.ToLower().Contains(term)) ||
                    (x.Descripcion != null && x.Descripcion.ToLower().Contains(term)) ||
                    x.Cod_Inte.ToString().Contains(term)
                );
            }

            var results = await query
                .OrderBy(x => x.Codigo)
                .Take(pageSize)
                .Select(x => new Instruci
                {
                    Id          = x.Id,
                    Cod_Inte    = x.Cod_Inte,
                    Codigo      = x.Codigo,
                    Descripcion = x.Descripcion,
                    Revision    = x.Revision,
                    Preparo     = x.Preparo,
                    FechaPrep   = x.FechaPrep,
                    Aprobo      = x.Aprobo,
                    FechaPro    = x.FechaPro,
                    Observaciones = x.Observaciones
                })
                .ToListAsync();

            return results;
        }

        /// <summary>
        /// Obtiene la primera instrucción que coincide con Cod_Inte
        /// GET /api/instruci/by-codinte/{codInte}
        /// </summary>
        [HttpGet("by-codinte/{codInte:int}")]
        public async Task<ActionResult<Instruci>> GetByCodInte(int codInte)
        {
            var item = await _context.Instrucis
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.Cod_Inte == codInte);
            if (item == null) return NotFound();
            return item;
        }

        /// <summary>
        /// Retorna el siguiente Cod_Inte disponible (max + 1)
        /// GET /api/instruci/next-cod-inte
        /// </summary>
        [HttpGet("next-cod-inte")]
        public async Task<ActionResult<int>> NextCodInte()
        {
            var max = await _context.Instrucis.MaxAsync(x => (int?)x.Cod_Inte) ?? 0;
            return max + 1;
        }

        /// <summary>
        /// Crea una nueva instrucción
        /// POST /api/instruci
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<Instruci>> Create([FromBody] Instruci item)
        {
            if (!ModelState.IsValid) return BadRequest(ModelState);
            item.AddRecord  = DateTime.Now;
            item.LastUpdate = DateTime.Now;
            _context.Instrucis.Add(item);
            await _context.SaveChangesAsync();
            return CreatedAtAction(nameof(GetById), new { id = item.Id }, item);
        }

        /// <summary>
        /// Obtiene una instrucción por Id
        /// GET /api/instruci/{id}
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<ActionResult<Instruci>> GetById(int id)
        {
            var item = await _context.Instrucis.AsNoTracking().FirstOrDefaultAsync(x => x.Id == id);
            if (item == null) return NotFound();
            return item;
        }

        /// <summary>
        /// Actualiza los campos editables de una instrucción (Preparo, FechaPrep, Aprobo, FechaPro)
        /// PUT /api/instruci/{id}
        /// </summary>
        [HttpPut("{id:int}")]
        public async Task<ActionResult> Update(int id, [FromBody] Instruci dto)
        {
            var entity = await _context.Instrucis.FirstOrDefaultAsync(x => x.Id == id);
            if (entity == null) return NotFound();

            entity.Revision      = dto.Revision;
            entity.Preparo       = dto.Preparo;
            entity.FechaPrep     = dto.FechaPrep;
            entity.Aprobo        = dto.Aprobo;
            entity.FechaPro      = dto.FechaPro;
            entity.Observaciones = dto.Observaciones;
            entity.LastUpdate    = DateTime.Now;

            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Obtiene las hojas de actualización de una instrucción por Cod_Inte, ordenadas por Orden
        /// GET /api/instruci/{codInte}/hojactu
        /// </summary>
        [HttpGet("{codInte:int}/hojactu")]
        public async Task<ActionResult<List<Hojactu>>> GetHojactu(int codInte)
        {
            var lista = await _context.Hojactus
                .AsNoTracking()
                .Where(x => x.Cod_Inte == codInte && x.Status == 0)
                .OrderBy(x => x.Orden)
                .ToListAsync();
            return lista;
        }

        /// <summary>
        /// Crea una nueva hoja de actualización
        /// POST /api/instruci/hojactu
        /// </summary>
        [HttpPost("hojactu")]
        public async Task<ActionResult<Hojactu>> CreateHojactu([FromBody] Hojactu item)
        {
            if (!ModelState.IsValid) return BadRequest(ModelState);
            item.AddRecord  = DateTime.Now;
            item.LastUpdate = DateTime.Now;
            item.CodiEmpr   = 0;
            _context.Hojactus.Add(item);
            await _context.SaveChangesAsync();
            return CreatedAtAction(nameof(GetHojactu), new { codInte = item.Cod_Inte }, item);
        }
    }
}
