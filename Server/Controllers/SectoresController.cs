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
    [Route("/api/sectores")]

    public class SectoresController : ControllerBase
    {
        private readonly DataContext _context;
        public SectoresController(DataContext context)
        {
            _context = context;
        }

        //[HttpGet]
        //public async Task<IActionResult> GetAsync()
        //{
        //    return Ok(await _context.Sectores.ToListAsync());
        //}

        //______________________________________________________________________
        [HttpGet]
        public async Task<ActionResult> Get([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Sectores
                //.Include(x => x.sector)
                .Where(x => x.DependenciaId == pagination.Id)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.sector!.ToLower().Contains(pagination.Filter.ToLower()));
            }

            return Ok(await queryable
                .OrderBy(x => x.sector)
                .Paginate(pagination)
                .ToListAsync());
        }
        //______________________________________________________________________
        [AllowAnonymous]
        [HttpGet("combo/{dependenciaId:int}")]
        public async Task<ActionResult> GetCombo(int dependenciaId)
        {
            return Ok(await _context.Sectores
                .Where(x => x.DependenciaId == dependenciaId)
                .ToListAsync());
        }
        //______________________________________________________________________
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Sectores
                .Where(x => x.Dependencia!.Id == pagination.Id)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.sector!.ToLower().Contains(pagination.Filter.ToLower()));
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }
        //______________________________________________________________________
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var city = await _context.Sectores.FirstOrDefaultAsync(x => x.Id == id);
            if (city == null)
            {
                return NotFound();
            }
            return Ok(city);
        }
        //______________________________________________________________________
        [HttpPost]
        public async Task<ActionResult> PostAsync(Sector sector)
        {
            try
            {
                _context.Add(sector); //con solo pasarle country ya interpreta a que tabla
                await _context.SaveChangesAsync();
                return Ok(sector);
            }

            catch (DbUpdateException dbUpdateException)
            {
                //se valida si se el nombre ya existe (dbUpdateException.InnerException arroja un error que contiene "duplicate"
                if (dbUpdateException.InnerException!.Message.Contains("duplicate")
                   || dbUpdateException.InnerException.Message.Contains("duplicada"))
                {
                    return BadRequest("Ya existe un sector con el mismo nombre.");
                }

                return BadRequest(dbUpdateException.Message);

            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }



        //______________________________________________________________________
        [HttpPut]
        public async Task<ActionResult> PutAsync(Sector sector)
        {
            try
            {   //no permite un duplicado (esto se declara en el datacontext en OnModelCreating)
                _context.Update(sector); //con solo pasarle country ya interpreta a que tabla
                await _context.SaveChangesAsync();
                return Ok(sector);
            }
            catch (DbUpdateException dbUpdateException)
            {   //se valida si se el nombre ya existe (dbUpdateException.InnerException arroja un error que contiene "duplicate"

                if (dbUpdateException.InnerException!.Message.Contains("duplicate")
                    || dbUpdateException.InnerException.Message.Contains("duplicada"))
                {
                    return BadRequest("Ya existe una sector con el mismo nombre.");
                }

                return BadRequest(dbUpdateException.Message);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }
        //______________________________________________________________________
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {

            var city = await _context.Sectores.FirstOrDefaultAsync(x => x.Id == id);
            if (city == null)
            {
                return NotFound();
            }
            _context.Remove(city);
            await _context.SaveChangesAsync();
            return NoContent();
        }
        //______________________________________________________________________
    }
}
