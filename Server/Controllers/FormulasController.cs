using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/formulas")]
    public class FormulasController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ISyncService _syncService;
        private readonly ILogger<FormulasController> _logger;

        public FormulasController(DataContext context, ISyncService syncService, ILogger<FormulasController> logger)
        {
            _context = context;
            _syncService = syncService;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene la lista de familias para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Formulas
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de fórmulas con filtros
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Formulas
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => 
                    (x.CodXConj != null && x.CodXConj.ToLower().Contains(pagination.Filter.ToLower())) ||
                    (x.CodXElem != null && x.CodXElem.ToLower().Contains(pagination.Filter.ToLower())));
            }
            return Ok(await queryable
                .OrderBy(x => x.CodIConj)
                .ThenBy(x => x.NuorIt)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene todos los componentes de un producto terminado
        /// </summary>//***** ESTO LO USA MASTER COMPONENTS ****
        /// 
        [HttpGet("conjunto/{codigoConjunto:int}")]
        public async Task<IActionResult> GetComponentesConjuntoAsync(int codigoConjunto)
        {
            try
            {
                var componentes = await _context.Formulas
             .Where(f => f.CodIConj == codigoConjunto)
             .OrderBy(f => f.NuorIt)
             .ToListAsync();

                return Ok(componentes);
            }
            catch (Exception ex)
            {
                throw;
            } 
         
        }

        /// <summary>
        /// Guarda o actualiza los componentes de un producto terminado
        /// </summary>
        [HttpPost("conjunto/{codigoConjunto:int}")]
        public async Task<IActionResult> SaveComponentesConjuntoAsync(int codigoConjunto, [FromBody] List<Formulas> componentes)
        {
            using var transaction = await _context.Database.BeginTransactionAsync(); // Iniciar transacción

            try
            {
                // Eliminar componentes existentes
                var componentesExistentes = await _context.Formulas
                    .Where(f => f.CodIConj == codigoConjunto)
                    .ToListAsync();

                _context.Formulas.RemoveRange(componentesExistentes);

                // Agregar nuevos componentes
                foreach (var componente in componentes)
                {
                    componente.Id = 0; // Asegurar que se cree como nuevo
                    componente.CodIConj = codigoConjunto; // Asignar el código correcto
                    _context.Formulas.Add(componente);
                }

                await _context.SaveChangesAsync(); // Guardar cambios
                await transaction.CommitAsync();  // Confirmar la transacción

                // Generar archivo de sincronización si está habilitado
                try
                {
                    await _syncService.GenerarArchivoSyncFormulasAsync(codigoConjunto);
                }
                catch (Exception syncEx)
                {
                    _logger.LogError(syncEx, $"Error al generar archivo de sincronización para Formulas CodIConj={codigoConjunto}");
                }

                return Ok();
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync(); // Deshacer los cambios si hay un error
                return BadRequest($"Error al guardar los componentes: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Formulas.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.CodXElem!.ToLower().Contains(pagination.Filter.ToLower()));
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene una Formulas por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var formula = await _context.Formulas
                .FirstOrDefaultAsync(x => x.Id == id);
            if (formula == null)
            {
                return NotFound();
            }
            return Ok(formula);
        }

        /// <summary>
        /// Crea una nueva formula
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Formulas formula)
        {
            try
            {
                // Verificar si ya existe una formula con el mismo código
                var existingFormula = await _context.Formulas
                    .FirstOrDefaultAsync(x => x.CodXElem == formula.CodXElem);

                if (existingFormula != null)
                {
                    return BadRequest($"Ya existe una formula con el código {formula.CodXElem}");
                }

                _context.Add(formula);
                await _context.SaveChangesAsync();
                return Ok(formula);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }
        /// <summary>
        /// Actualiza una formula existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Formulas formula)
        {
            try
            {
                // Verificar si ya existe una formula con el mismo código (excluyendo la actual)
                var existingFormula = await _context.Formulas
                    .FirstOrDefaultAsync(x => x.CodXElem == formula.CodXElem && x.Id != formula.Id);

                if (existingFormula != null)
                {
                    return BadRequest($"Ya existe una formula con el código {formula.CodXElem}");
                }

                _context.Update(formula);
                await _context.SaveChangesAsync();
                return Ok(formula);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Obtiene la fórmula de componentes de un producto con stock y cálculos adicionales
        /// </summary>
        [HttpGet("componentes/{codigoInternoProducto:int}")]
        public async Task<IActionResult> GetComponentesConStockAsync(int codigoInternoProducto)
        {
            try
            {
                // Obtener los componentes de la fórmula
                var componentes = await _context.Formulas
                    .Where(f => f.CodIConj == codigoInternoProducto)
                    .OrderBy(f => f.NuorIt)
                    .ToListAsync();

                if (componentes == null || !componentes.Any())
                {
                    return NotFound($"No se encontró fórmula para el producto con código interno {codigoInternoProducto}");
                }

                // Obtener todos los códigos de elementos para verificar repetidos
                var codigosElementos = componentes.Select(c => c.CodiElem).ToList();
                var codigosRepetidos = codigosElementos.GroupBy(c => c).Where(g => g.Count() > 1).Select(g => g.Key).ToList();

                // Crear lista de DTOs con información adicional
                var resultado = new List<FormulaComponenteDTO>();

                foreach (var componente in componentes)
                {
                    if (componente.CodiElem == null) continue;

                    // Buscar el master correspondiente para obtener la descripción
                    var master = await _context.masters
                        .FirstOrDefaultAsync(m => m.Codint == componente.CodiElem);

                  

                    // Crear el DTO con la información
                    var componenteDTO = new FormulaComponenteDTO
                    {
                        CodIElem = componente.CodiElem ?? 0,
                        CodXElem = componente.CodXElem ?? string.Empty,
                        Descripcion = master?.Descripcion ?? "Sin descripción",
                        UsoBruto = componente.UsoBruto ?? 0,
                        EsRepetido = codigosRepetidos.Contains(componente.CodiElem),
                        NecesidadTotal = 0, // Se calculará en el frontend con la cantidad a fabricar
                        CantidadAUtilizar = 0 // Se inicializará en el frontend
                    };

                    resultado.Add(componenteDTO);
                }
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener los componentes: {ex.Message}");
            }
        }

        /// <summary>
        /// Elimina una formula (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var formula = await _context.Formulas.FirstOrDefaultAsync(x => x.Id == id);
            if (formula == null)
            {
                return NotFound();
            }
            _context.Remove(formula);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca una formula como inactiva (borrado lógico)
        /// </summary>
        //[HttpPut("baja/{id:int}")]
        //public async Task<IActionResult> FamiliaStatusMenosUnoAsync(int id)
        //{
        //    var formula = await _context.Formulas.FirstOrDefaultAsync(x => x.Id == id);
        //    if (formula == null)
        //    {
        //        return NotFound();
        //    }
        //    formula.Status = -1;
        //    _context.Update(formula);
        //    await _context.SaveChangesAsync();
        //    return NoContent();
        //}
        //______________________________________________________________________
        //[HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        //public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        //{
        //    var familia = await _context.Familias.FirstOrDefaultAsync(x => x.Id == id);
        //    if (familia == null)
        //    {
        //        return NotFound();
        //    }

        //    // Cambiar el campo status 
        //    familia.Status = status;

        //    // Guardar los cambios en la base de datos
        //    await _context.SaveChangesAsync();

        //    return NoContent();
        //}
    }
}
