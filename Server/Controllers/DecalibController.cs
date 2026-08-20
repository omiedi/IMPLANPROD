using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class DecalibController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<DecalibController> _logger;

        public DecalibController(DataContext context, ILogger<DecalibController> logger)
        {
            _context = context;
            _logger = logger;
        }

        [HttpGet("by-recalib/{recalibId:int}")]
        public async Task<ActionResult<List<Decalib>>> GetByRecalibId(int recalibId)
        {
            try
            {
                var recalibRef = await _context.Recalibs
                    .AsNoTracking()
                    .Where(r => r.Id == recalibId)
                    .Select(r => new { r.Id, r.RegCalibrNro })
                    .FirstOrDefaultAsync();

                if (recalibRef == null)
                {
                    return Ok(new List<Decalib>());
                }

                var registros = await _context.Decalibs
                    .Where(d => d.IdReCalib == recalibRef.Id
                             || (recalibRef.RegCalibrNro.HasValue && d.IdReCalib == recalibRef.RegCalibrNro.Value))
                    .OrderBy(d => d.Id)
                    .ToListAsync();

                return Ok(registros);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener DECALIB para RECALIB.Id {RecalibId}", recalibId);
                return StatusCode(500, new { error = "Error al obtener datos de DECALIB" });
            }
        }

        [HttpPut("{id:int}")]
        public async Task<ActionResult<Decalib>> Update(int id, [FromBody] Decalib model)
        {
            try
            {
                var entity = await _context.Decalibs.FirstOrDefaultAsync(d => d.Id == id);
                if (entity == null)
                {
                    return NotFound(new { error = "Registro DECALIB no encontrado" });
                }

                entity.Patron = model.Patron;
                entity.Medida = model.Medida;
                entity.X1 = model.X1;
                entity.X2 = model.X2;
                entity.X3 = model.X3;
                entity.X4 = model.X4;
                entity.X5 = model.X5;
                entity.EProm = model.EProm;
                entity.IMe = model.IMe;
                entity.LastUpdate = DateTime.Now;
                entity.CodiEmprNet ??= 0;

                await _context.SaveChangesAsync();

                return Ok(entity);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar DECALIB.Id {Id}", id);
                return StatusCode(500, new { error = "Error al actualizar DECALIB" });
            }
        }

        [HttpPost("by-recalib/{recalibId:int}")]
        public async Task<ActionResult<Decalib>> CreateByRecalibId(int recalibId, [FromBody] Decalib model)
        {
            try
            {
                var recalibRef = await _context.Recalibs
                    .AsNoTracking()
                    .Where(r => r.Id == recalibId)
                    .Select(r => new { r.Id })
                    .FirstOrDefaultAsync();

                if (recalibRef == null)
                {
                    return BadRequest(new { error = "No se encontró RECALIB válido para asociar DECALIB" });
                }

                var entity = new Decalib
                {
                    Patron = model.Patron,
                    Medida = model.Medida,
                    X1 = model.X1,
                    X2 = model.X2,
                    X3 = model.X3,
                    X4 = model.X4,
                    X5 = model.X5,
                    EProm = model.EProm,
                    IMe = model.IMe,
                    IdReCalib = recalibRef.Id,
                    Status = 0,
                    AddRecord = DateTime.Now,
                    LastUpdate = DateTime.Now,
                    CodiEmprNet = model.CodiEmprNet ?? 0,
                    NumUsuar = model.NumUsuar
                };

                _context.Decalibs.Add(entity);
                await _context.SaveChangesAsync();

                return Ok(entity);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear DECALIB para RECALIB.Id {RecalibId}", recalibId);
                return StatusCode(500, new { error = "Error al crear DECALIB" });
            }
        }

        [HttpDelete("{id:int}")]
        public async Task<ActionResult> Delete(int id)
        {
            try
            {
                var entity = await _context.Decalibs.FirstOrDefaultAsync(d => d.Id == id);
                if (entity == null)
                {
                    return NotFound(new { error = "Registro DECALIB no encontrado" });
                }

                _context.Decalibs.Remove(entity);
                await _context.SaveChangesAsync();

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al eliminar DECALIB.Id {Id}", id);
                return StatusCode(500, new { error = "Error al eliminar DECALIB" });
            }
        }
    }
}
