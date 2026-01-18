using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class RepurepaController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<RepurepaController> _logger;

        public RepurepaController(DataContext context, ILogger<RepurepaController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene los repuestos de una reparación (NUMREPA) con datos del MASTER.
        /// 
        /// Equivalente VB6:
        /// SELECT * FROM REPUREPA WHERE NUMREPA = ... AND ORIGEN=0
        /// 
        /// Notas:
        /// - Se excluyen items marcados como borrados (MarBorra <> 0)
        /// - Se devuelve DTO con Código/Descripción/U.M. resueltos desde MASTER
        /// </summary>
        [HttpGet("reparacion/{numRepa:int}")]
        public async Task<ActionResult<List<RepurepaDTO>>> GetByReparacion([FromRoute] int numRepa, [FromQuery] int origen = 0)
        {
            try
            {
                var query =
                    from r in _context.Repurepas
                    join m in _context.masters on r.Cod_Inte equals m.Codint into masterJoin
                    from m in masterJoin.DefaultIfEmpty()
                    where r.NumRepa == numRepa
                          && (r.Origen ?? 0) == origen
                          && (r.MarBorra ?? 0) == 0
                    orderby r.Nuord_Item, r.Id
                    select new RepurepaDTO
                    {
                        Id = r.Id,
                        NumRepa = r.NumRepa,
                        Cod_Inte = r.Cod_Inte,
                        Cant_Sol = r.Cant_Sol ?? 0m,
                        CodigoExterno = m.Codigo,
                        Descripcion = m.Descripcion,
                        UnidadMedida = m.Umedida,
                        Origen = r.Origen,
                        MarBorra = r.MarBorra,
                        Nuord_Item = r.Nuord_Item
                    };

                var items = await query.ToListAsync();
                return Ok(items);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener repuestos de la reparación {NumRepa}", numRepa);
                return StatusCode(500, $"Error al obtener repuestos. Detalle: {ex.Message}");
            }
        }

        /// <summary>
        /// Crea un ítem de repuesto para una reparación.
        /// 
        /// Objetivo:
        /// - Guardado independiente (cada alta es una transacción individual)
        /// - Asignar Nuord_Item secuencial dentro de la reparación
        /// - Dejar valores monetarios en 0 si no vienen (se podrán completar luego)
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<Repurepa>> PostAsync([FromBody] Repurepa item)
        {
            try
            {
                if (item.NumRepa <= 0)
                {
                    return BadRequest("NumRepa inválido");
                }

                if (item.Cod_Inte <= 0)
                {
                    return BadRequest("Cod_Inte inválido");
                }

                // Normalizar defaults
                item.Origen ??= 0;
                item.MarBorra ??= 0;

                // Asignar número de orden secuencial (VB6 usaba JJ& y refrescaba la lista)
                var maxNuord = await _context.Repurepas
                    .Where(x => x.NumRepa == item.NumRepa && (x.Origen ?? 0) == (item.Origen ?? 0))
                    .MaxAsync(x => (short?)x.Nuord_Item) ?? (short)0;

                item.Nuord_Item = (short)(maxNuord + 1);

                // Si no viene cantidad, usar 1 (para evitar registros "vacíos")
                if (!item.Cant_Sol.HasValue || item.Cant_Sol.Value <= 0)
                {
                    item.Cant_Sol = 1m;
                }

                // Valores monetarios: mantener 0 si no vienen
                item.PrUn_Neto ??= 0m;
                item.PORDESC ??= 0m;

                // Recalcular importe neto simple
                item.Impo_Neto = Math.Round((item.Cant_Sol ?? 0m) * (item.PrUn_Neto ?? 0m), 4);

                _context.Repurepas.Add(item);
                await _context.SaveChangesAsync();

                return Ok(item);
            }
            catch (DbUpdateException ex)
            {
                _logger.LogError(ex, "Error de BD al crear repuesto (REPUREPA) para NumRepa={NumRepa}", item.NumRepa);
                return StatusCode(500, $"Error de base de datos al crear el repuesto. Detalle: {ex.InnerException?.Message ?? ex.Message}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear repuesto (REPUREPA) para NumRepa={NumRepa}", item.NumRepa);
                return StatusCode(500, $"Error al crear el repuesto. Detalle: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza la cantidad (CANT_SOL) de un ítem.
        /// 
        /// Nota:
        /// - El frontend edita directamente el input y llama a este endpoint.
        /// - Recalcula Impo_Neto en base a Cant_Sol * PrUn_Neto.
        /// </summary>
        [HttpPut("{id:int}/cantidad")]
        public async Task<ActionResult> UpdateCantidadAsync([FromRoute] int id, [FromBody] decimal cantidad)
        {
            try
            {
                var item = await _context.Repurepas.FirstOrDefaultAsync(x => x.Id == id);
                if (item == null)
                {
                    return NotFound("Ítem no encontrado");
                }

                if (cantidad <= 0)
                {
                    return BadRequest("La cantidad debe ser mayor a 0");
                }

                // Actualizar cantidad (formato con 1 decimal se maneja en UI, acá guardamos con 4)
                item.Cant_Sol = Math.Round(cantidad, 4);

                // Recalcular importe neto
                item.Impo_Neto = Math.Round((item.Cant_Sol ?? 0m) * (item.PrUn_Neto ?? 0m), 4);

                await _context.SaveChangesAsync();
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar cantidad REPUREPA Id={Id}", id);
                return StatusCode(500, $"Error al actualizar cantidad. Detalle: {ex.Message}");
            }
        }

        /// <summary>
        /// Borrado lógico del ítem: MarBorra = 1.
        /// 
        /// Nota:
        /// - Se usa para permitir confirmación desde UI y mantener trazabilidad.
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<ActionResult> DeleteAsync([FromRoute] int id)
        {
            try
            {
                var item = await _context.Repurepas.FirstOrDefaultAsync(x => x.Id == id);
                if (item == null)
                {
                    return NotFound("Ítem no encontrado");
                }

                item.MarBorra = 1;
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al borrar (lógico) REPUREPA Id={Id}", id);
                return StatusCode(500, $"Error al borrar el ítem. Detalle: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene repuestos de una reparación para orden de reparación (ORIGEN=1)
        /// Equivalente VB6: SELECT * FROM REPUREPA WHERE NumRepa = {NPX} AND ORIGEN=1 ORDER BY NuOrd_Item ASC
        /// </summary>
        [HttpGet("por-reparacion/{numRepa:int}")]
        public async Task<ActionResult<List<Repurepa>>> GetPorReparacionAsync([FromRoute] int numRepa)
        {
            try
            {
                var repuestos = await _context.Repurepas
                    .Where(r => r.NumRepa == numRepa && r.Origen == 1 && (r.MarBorra ?? 0) == 0)
                    .OrderBy(r => r.Nuord_Item)
                    .ToListAsync();

                _logger.LogInformation("Repuestos cargados para NumRepa={NumRepa}, ORIGEN=1: {Count}", numRepa, repuestos.Count);
                
                return Ok(repuestos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener repuestos para NumRepa={NumRepa}", numRepa);
                return StatusCode(500, $"Error al obtener repuestos: {ex.Message}");
            }
        }

        /// <summary>
        /// Borra todos los repuestos de una reparación con ORIGEN específico
        /// y guarda nuevos repuestos desde el presupuesto
        /// </summary>
        [HttpPost("guardar-presupuesto")]
        public async Task<ActionResult> GuardarRepuestosPresupuesto([FromBody] GuardarRepuestosPresupuestoRequest request)
        {
            try
            {
                if (request.NumRepa <= 0)
                {
                    return BadRequest("NumRepa inválido");
                }

                // 1. Borrar repuestos existentes con ORIGEN=1 (comercial)
                var repuestosExistentes = await _context.Repurepas
                    .Where(r => r.NumRepa == request.NumRepa && r.Origen == 1)
                    .ToListAsync();

                if (repuestosExistentes.Any())
                {
                    _context.Repurepas.RemoveRange(repuestosExistentes);
                }

                // 2. Guardar nuevos repuestos
                if (request.Repuestos != null && request.Repuestos.Any())
                {
                    short contador = 1;
                    foreach (var repuesto in request.Repuestos)
                    {
                        var nuevoRepuesto = new Repurepa
                        {
                            NumRepa = request.NumRepa,
                            Cod_Inte = repuesto.CodigoInterno,
                            Cant_Sol = repuesto.Cantidad,
                            PrUn_Neto = 0,
                            MarBorra = 0,
                            Impo_Neto = 0,
                            Nuord_Item = contador,
                            IdenLote = "",
                            Origen = 1,
                            CodiEmpr = 0
                        };

                        _context.Repurepas.Add(nuevoRepuesto);
                        contador++;
                    }
                }

                await _context.SaveChangesAsync();

                _logger.LogInformation("Repuestos de presupuesto guardados para NumRepa={NumRepa}", request.NumRepa);
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al guardar repuestos de presupuesto NumRepa={NumRepa}", request.NumRepa);
                return StatusCode(500, $"Error al guardar repuestos. Detalle: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Request para guardar repuestos del presupuesto
    /// </summary>
    public class GuardarRepuestosPresupuestoRequest
    {
        public int NumRepa { get; set; }
        public List<RepuestoPresupuestoItem>? Repuestos { get; set; }
    }

    /// <summary>
    /// Item de repuesto para guardar
    /// </summary>
    public class RepuestoPresupuestoItem
    {
        public int CodigoInterno { get; set; }
        public decimal Cantidad { get; set; }
    }
}
