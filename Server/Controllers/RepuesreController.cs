using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme)]
    [Route("api/[controller]")]
    public class RepuesreController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<RepuesreController> _logger;

        public RepuesreController(DataContext context, ILogger<RepuesreController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene los repuestos de una reparación con información del Master
        /// </summary>
        [HttpGet("reparacion/{codigoReparacion}")]
        public async Task<ActionResult<List<RepuestoReparacionDTO>>> GetRepuestosPorReparacion(string codigoReparacion)
        {
            try
            {
                var repuestos = await _context.Repuesres
                    .Where(r => r.CodigoReparacion == codigoReparacion && r.Status >= 0)
                    .Join(_context.masters,
                        repuesto => repuesto.CodigoInterno,
                        master => master.Codint,
                        (repuesto, master) => new RepuestoReparacionDTO
                        {
                            Id = repuesto.Id,
                            CodigoReparacion = repuesto.CodigoReparacion,
                            CodigoInterno = repuesto.CodigoInterno,
                            CodigoExterno = master.Codigo,
                            UnidadMedida = master.Umedida,
                            Descripcion = master.Descripcion,
                            Uso = repuesto.Uso
                        })
                    .OrderBy(r => r.CodigoInterno)
                    .ToListAsync();

                return Ok(repuestos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener repuestos de la reparación {CodigoReparacion}", codigoReparacion);
                return StatusCode(500, "Error al obtener los repuestos");
            }
        }

        /// <summary>
        /// Agrega un nuevo repuesto a una reparación
        /// </summary>
        [HttpPost("agregar")]
        public async Task<ActionResult<RepuestoReparacionDTO>> AgregarRepuesto([FromBody] AgregarRepuestoRequest request)
        {
            try
            {
                // Verificar que el producto existe en Master
                var producto = await _context.masters
                    .Where(m => m.Codint == request.CodigoInterno && m.Status >= 0)
                    .FirstOrDefaultAsync();

                if (producto == null)
                {
                    return BadRequest("El producto no existe o está inactivo");
                }

                // Verificar que no exista ya este repuesto para esta reparación
                var existente = await _context.Repuesres
                    .Where(r => r.CodigoReparacion == request.CodigoReparacion 
                             && r.CodigoInterno == request.CodigoInterno
                             && r.Status >= 0)
                    .FirstOrDefaultAsync();

                if (existente != null)
                {
                    return BadRequest("Este repuesto ya está agregado a la reparación");
                }

                // Crear nuevo registro
                var nuevoRepuesto = new Repuesre
                {
                    CodigoReparacion = request.CodigoReparacion,
                    CodigoInterno = request.CodigoInterno,
                    Uso = request.Uso,
                    Status = 0
                };

                _context.Repuesres.Add(nuevoRepuesto);
                await _context.SaveChangesAsync();

                // Retornar el DTO completo con información del Master
                var resultado = new RepuestoReparacionDTO
                {
                    Id = nuevoRepuesto.Id,
                    CodigoReparacion = nuevoRepuesto.CodigoReparacion,
                    CodigoInterno = nuevoRepuesto.CodigoInterno,
                    CodigoExterno = producto.Codigo,
                    UnidadMedida = producto.Umedida,
                    Descripcion = producto.Descripcion,
                    Uso = nuevoRepuesto.Uso
                };

                _logger.LogInformation("Repuesto agregado: {CodigoReparacion} - {CodigoInterno}", 
                    request.CodigoReparacion, request.CodigoInterno);

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al agregar repuesto");
                return StatusCode(500, "Error al agregar el repuesto");
            }
        }

        /// <summary>
        /// Elimina un repuesto de una reparación (marca como inactivo)
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<ActionResult> EliminarRepuesto(int id)
        {
            try
            {
                var repuesto = await _context.Repuesres.FindAsync(id);

                if (repuesto == null)
                {
                    return NotFound("Repuesto no encontrado");
                }

                // Marcar como inactivo en lugar de eliminar físicamente
                repuesto.Status = -1;
                await _context.SaveChangesAsync();

                _logger.LogInformation("Repuesto eliminado: ID {Id} - {CodigoReparacion}", 
                    id, repuesto.CodigoReparacion);

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al eliminar repuesto {Id}", id);
                return StatusCode(500, "Error al eliminar el repuesto");
            }
        }

        /// <summary>
        /// Actualiza la cantidad de uso de un repuesto
        /// </summary>
        [HttpPut("{id}/uso")]
        public async Task<ActionResult> ActualizarUso(int id, [FromBody] decimal nuevoUso)
        {
            try
            {
                var repuesto = await _context.Repuesres.FindAsync(id);

                if (repuesto == null)
                {
                    return NotFound("Repuesto no encontrado");
                }

                repuesto.Uso = nuevoUso;
                await _context.SaveChangesAsync();

                _logger.LogInformation("Uso actualizado: ID {Id} - Nuevo uso: {Uso}", id, nuevoUso);

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar uso del repuesto {Id}", id);
                return StatusCode(500, "Error al actualizar el uso");
            }
        }

        /// <summary>
        /// Obtiene los repuestos para presupuesto con cálculos de costos
        /// Aplica el coeficiente de presupuesto (COEFPREA) al costo unitario
        /// </summary>
        [HttpGet("presupuesto/{codigoReparacion}")]
        public async Task<ActionResult<List<RepuestoPresupuestoDTO>>> GetRepuestosParaPresupuesto(string codigoReparacion, [FromQuery] decimal coeficiente = 1)
        {
            try
            {
                var repuestos = await _context.Repuesres
                    .Where(r => r.CodigoReparacion == codigoReparacion && r.Status >= 0)
                    .Join(_context.masters,
                        repuesto => repuesto.CodigoInterno,
                        master => master.Codint,
                        (repuesto, master) => new RepuestoPresupuestoDTO
                        {
                            CodigoInterno = repuesto.CodigoInterno,
                            Codigo = master.Codigo,
                            Descripcion = master.Descripcion,
                            Uso = repuesto.Uso,
                            CostoUnitarioBase = master.Cosuni,
                            CostoUnitario = (master.Cosuni ?? 0) * coeficiente,
                            Importe = repuesto.Uso * ((master.Cosuni ?? 0) * coeficiente)
                        })
                    .OrderBy(r => r.CodigoInterno)
                    .ToListAsync();

                return Ok(repuestos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener repuestos para presupuesto {CodigoReparacion}", codigoReparacion);
                return StatusCode(500, "Error al obtener los repuestos para presupuesto");
            }
        }
    }
}
