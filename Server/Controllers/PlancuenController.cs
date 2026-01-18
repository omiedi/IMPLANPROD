using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/plancuen")]
    public class PlancuenController : ControllerBase
    {
        private readonly DataContext _context;

        public PlancuenController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de cuentas para el combo (solo activas)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Plancuens
                .Where(d => d.Statcue >= 0)
                .OrderBy(x => x.CodExCue)
                .ToListAsync());
        }
        /// <summary>
        /// Obtiene la lista paginada de cuentas con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="estado">Estado de la cuenta (1=activo, -1=inactivo)</param>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, int estado = 1)
        {
            var queryable = _context.Plancuens.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.CodExCue!.Contains(pagination.Filter) ||
                                                x.Denocue!.ToLower().Contains(pagination.Filter.ToLower()));
            }

            // Filtrar por estado
            if (estado == 1)
            {
                queryable = queryable.Where(x => x.Statcue == estado || x.Statcue == null); // Activos
            }
            else
            {
                queryable = queryable.Where(x => x.Statcue == estado); // Inactivos
            }

            return Ok(await queryable
                .OrderBy(x => x.CodExCue)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 1)
        {
            var queryable = _context.Plancuens.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.CodExCue!.Contains(pagination.Filter) ||
                                                x.Denocue!.ToLower().Contains(pagination.Filter.ToLower()));
            }

            // Filtrar por estado
            if (estado == 1)
            {
                queryable = queryable.Where(x => x.Statcue == estado || x.Statcue == null);
            }
            else
            {
                queryable = queryable.Where(x => x.Statcue == estado);
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene una cuenta por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var cuenta = await _context.Plancuens
                .FirstOrDefaultAsync(x => x.Id == id);
            if (cuenta == null)
            {
                return NotFound();
            }
            return Ok(cuenta);
        }
        /// <summary>
        /// Obtiene el siguiente código interno disponible
        /// </summary>
        [HttpGet("siguiente-codigo")]
        public async Task<IActionResult> GetSiguienteCodigoAsync()
        {
            var maxCodigo = await _context.Plancuens
                .Where(x => x.CodInCue.HasValue)
                .MaxAsync(x => (int?)x.CodInCue) ?? 0;

            return Ok(maxCodigo + 1);
        }

        /// <summary>
        /// Crea una nueva cuenta
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Plancuen cuenta)
        {
            try
            {
                // Formatear CodExCue a 2 decimales si es numérico
                if (!string.IsNullOrEmpty(cuenta.CodExCue) && decimal.TryParse(cuenta.CodExCue, out decimal codigo))
                {
                    cuenta.CodExCue = codigo.ToString("F2");
                }

                // Verificar si ya existe una cuenta con el mismo código externo
                var existingCuenta = await _context.Plancuens
                    .FirstOrDefaultAsync(x => x.CodExCue == cuenta.CodExCue);

                if (existingCuenta != null)
                {
                    return BadRequest($"Ya existe una cuenta con el código {cuenta.CodExCue}");
                }

                // Establecer valores por defecto
                if (!cuenta.CodInCue.HasValue)
                {
                    var maxCodigo = await _context.Plancuens
                        .Where(x => x.CodInCue.HasValue)
                        .MaxAsync(x => (int?)x.CodInCue) ?? 0;
                    cuenta.CodInCue = maxCodigo + 1;
                }

                if (string.IsNullOrEmpty(cuenta.TipoCue))
                {
                    cuenta.TipoCue = "debe";
                }

                //cuenta.AddRecord = DateTime.Now;
                cuenta.Statcue = 1; // Activo por defecto

                _context.Add(cuenta);
                await _context.SaveChangesAsync();
                return Ok(cuenta);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }
        /// <summary>
        /// Actualiza una cuenta existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Plancuen cuenta)
        {
            try
            {
                // Formatear CodExCue a 2 decimales si es numérico
                if (!string.IsNullOrEmpty(cuenta.CodExCue) && decimal.TryParse(cuenta.CodExCue, out decimal codigo))
                {
                    cuenta.CodExCue = codigo.ToString("F2");
                }

                // Verificar si ya existe una cuenta con el mismo código externo (excluyendo la actual)
                var existingCuenta = await _context.Plancuens
                    .FirstOrDefaultAsync(x => x.CodExCue == cuenta.CodExCue && x.Id != cuenta.Id);

                if (existingCuenta != null)
                {
                    return BadRequest($"Ya existe una cuenta con el código {cuenta.CodExCue}");
                }

                //cuenta.LastUpdate = DateTime.Now;

                _context.Update(cuenta);
                await _context.SaveChangesAsync();
                return Ok(cuenta);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }
        /// <summary>
        /// Elimina una cuenta (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var cuenta = await _context.Plancuens.FirstOrDefaultAsync(x => x.Id == id);
            if (cuenta == null)
            {
                return NotFound();
            }
            _context.Remove(cuenta);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca una cuenta como inactiva(borrado lógico)
        /// </summary>
        [HttpPut("baja/{id:int}")]
        public async Task<IActionResult> CuentaStatusMenosUnoAsync(int id)
        {
            var cuenta = await _context.Plancuens.FirstOrDefaultAsync(x => x.Id == id);
            if (cuenta == null)
            {
                return NotFound();
            }
            cuenta.Statcue = -1;
            //cuenta.LastUpdate = DateTime.Now;
            _context.Update(cuenta);
            await _context.SaveChangesAsync();
            return NoContent();
        }
        /// <summary>
        /// Marca una cuenta como inactiva(borrado lógico)
        /// </summary>
        [HttpPut("Restaura/{id:int}")]
        public async Task<IActionResult> CuentaStatusAUnoAsync(int id)
        {
            var cuenta = await _context.Plancuens.FirstOrDefaultAsync(x => x.Id == id);
            if (cuenta == null)
            {
                return NotFound();
            }
            cuenta.Statcue = 1;
            //cuenta.LastUpdate = DateTime.Now;
            _context.Update(cuenta);
            await _context.SaveChangesAsync();
            return NoContent();
        }
        /////// <summary>
        /////// Cambia el estado de una cuenta
        /////// </summary>
        //[HttpPut("cambiaStatus/{id:int}/{status:short}")]
        //public async Task<IActionResult> MarcaStatusAsync(int id, short status )
        //{
        //    var cuenta = await _context.Plancuens.FirstOrDefaultAsync(x => x.Id == id);
        //    if (cuenta == null)
        //    {
        //        return NotFound();
        //    }

        //    // Cambiar el campo status 
        //    cuenta.Statcue = status;
        //    //cuenta.LastUpdate = DateTime.Now;

        //    // Marcar la entidad como modificada y guardar los cambios
        //    //_context.Update(cuenta);
        //    await _context.SaveChangesAsync();

        //    return NoContent();
        //}


    }
}
