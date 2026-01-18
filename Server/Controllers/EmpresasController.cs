using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Server.Helpers;


namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/empresas")]
    public class EmpresasController : ControllerBase
    {
        private readonly DataContext _context;

        public EmpresasController(DataContext context)
        {
            _context = context;
        }


        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Empresas.ToListAsync());
        }

        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Empresas
                .Include(x => x.Dependencias)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.razon.ToLower().Contains(pagination.Filter.ToLower()));
            }
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.status == estado || x.status == null);//por defecto es = 0 
            }
            else
            {
                queryable = queryable.Where(x => x.status == estado);
            }
            return Ok(await queryable
            .OrderBy(x => x.razon)
            .Paginate(pagination)
            .ToListAsync());
        }

        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Empresas.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.razon.ToLower().Contains(pagination.Filter.ToLower()));
            }
            if (estado == 0)
            {
                queryable = queryable.Where(x => x.status == estado || x.status == null);//por defecto es = 0 
            }
            else
            {
                queryable = queryable.Where(x => x.status == estado);
            }
            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        [HttpGet("full")]
        public async Task<IActionResult> GetFullAsync()
        {
            return Ok(await _context.Empresas
                .Include(x => x.Dependencias!)
                .ThenInclude(x => x.Sectores)
                .ToListAsync());
        }
        //______________________________________________________________________

        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var empresa = await _context.Empresas
                .Include(x => x.Dependencias!)
                .ThenInclude(x => x.Sectores)
                .FirstOrDefaultAsync(x => x.Id == id);
            if (empresa == null)
            {
                return NotFound();
            }

            return Ok(empresa);
        }

        [HttpPost]
        public async Task<ActionResult> PostAsync(Empresa empresa)
        {
            try
            {
                _context.Add(empresa);
                await _context.SaveChangesAsync();
                return Ok(empresa);
            }
            catch (DbUpdateException dbUpdateException)
            {
                if (dbUpdateException.InnerException!.Message.Contains("duplicate")
                  || dbUpdateException.InnerException.Message.Contains("duplicada"))
                {
                    return BadRequest("Ya existe una Empresa con el mismo nombre.");
                }

                return BadRequest(dbUpdateException.Message);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        [HttpPut]
        public async Task<ActionResult> PutAsync(Empresa empresa)
        {
            try
            {
                _context.Update(empresa);
                await _context.SaveChangesAsync();
                return Ok(empresa);
            }
            catch (DbUpdateException dbUpdateException)
            {
                if (dbUpdateException.InnerException!.Message.Contains("duplicate")
                  || dbUpdateException.InnerException.Message.Contains("duplicada"))
                {
                    return BadRequest("Ya existe una empresa con el mismo nombre.");
                }

                return BadRequest(dbUpdateException.Message);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var empresa = await _context.Empresas.FirstOrDefaultAsync(x => x.Id == id);
            if (empresa == null)
            {
                return NotFound();
            }

            _context.Remove(empresa);
            await _context.SaveChangesAsync();
            return NoContent();
        }
    }
}
