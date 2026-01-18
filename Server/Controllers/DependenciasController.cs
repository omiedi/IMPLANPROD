using IMPLANPROD.Server.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Authorization;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;



namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/dependencias")]
    public class DependenciasController : ControllerBase
    {
        private readonly DataContext _context;
        public DependenciasController(DataContext context)
        {
            _context = context;
        }
        [AllowAnonymous]
        [HttpGet("combo/{empresaId:int}")]
        public async Task<ActionResult> GetCombo(int empresaId)
        {
            return Ok(await _context.Dependencias
                .Where(x => x.EmpresaId == empresaId)
                .ToListAsync());
        }



        //[HttpGet]
        //public async Task<IActionResult> GetAsync()
        //{
        //    return Ok(await _context.Dependencias.ToListAsync());
        //}
        //______________________________________________________________________
        [HttpGet]
        public async Task<ActionResult> Get([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Dependencias
                .Include(x => x.Sectores)
                .Where(x => x.Empresa!.Id == pagination.Id)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.dependencia!.ToLower().Contains(pagination.Filter.ToLower()));
            }

            return Ok(await queryable
                .OrderBy(x => x.dependencia)
                .Paginate(pagination)
                .ToListAsync());
        }

        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Dependencias
                .Where(x => x.Empresa!.Id == pagination.Id)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.dependencia!.ToLower().Contains(pagination.Filter.ToLower()));
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var dependencia = await _context.Dependencias
                .Include(x => x.Sectores!)
                .FirstOrDefaultAsync(x => x.Id == id);
            if (dependencia == null)
            {
                return NotFound();
            }
            return Ok(dependencia);
        }
        //______________________________________________________________________
        [HttpPost]
        public async Task<ActionResult> PostAsync(Dependencia dependencias)
        {
            try

            {
                _context.Add(dependencias); //con solo pasarle maquina ya interpreta a que tabla
                await _context.SaveChangesAsync();
                return Ok(dependencias);
            }

            catch (DbUpdateException dbUpdateException)
            {
                //se valida si se el nombre ya existe (dbUpdateException.InnerException arroja un error que contiene "duplicada"
                if (dbUpdateException.InnerException!.Message.Contains("duplicate")
                   || dbUpdateException.InnerException.Message.Contains("duplicada"))
                {
                    return BadRequest("Ya existe una filial con el mismo nombre.");
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
        public async Task<ActionResult> PutAsync(Dependencia dependencia)
        {
            try
            {   //no permite un duplicado (esto se declara en el datacontext en OnModelCreating)

                _context.Update(dependencia); //con solo pasarle country ya interpreta a que tabla
                await _context.SaveChangesAsync();
                return Ok(dependencia);
            }
            catch (DbUpdateException dbUpdateException)
            {   //se valida si se el nombre ya existe (dbUpdateException.InnerException arroja un error que contiene "duplicada"
                if (dbUpdateException.InnerException!.Message.Contains("duplicate")
                                   || dbUpdateException.InnerException.Message.Contains("duplicada"))
                {
                    return BadRequest("Ya existe una filial con el mismo nombre.");
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
            var dependencia = await _context.Dependencias.FirstOrDefaultAsync(x => x.Id == id);
            if (dependencia == null)
            {
                return NotFound();
            }
            _context.Remove(dependencia);
            await _context.SaveChangesAsync();
            return NoContent();
        }
        //______________________________________________________________________
    }
}

