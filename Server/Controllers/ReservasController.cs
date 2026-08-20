using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para manejar las operaciones de reservas de órdenes de fabricación
    /// </summary>
    [ApiController]
    [Route("/api/reservas")]
    public class ReservasController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ReservasService _reservasService;
        private readonly IUsuarioFlexoftService _UsuarioFlexoftService;
        private readonly ICodiempresa _Codiempresa;


        /// <summary>
        /// Constructor del controlador de reservas
        /// </summary>
        /// <param name="context">Contexto de base de datos</param>
        /// <param name="reservasService">Servicio de reservas</param>
        public ReservasController(DataContext context, ReservasService reservasService, IUsuarioFlexoftService UsuarioFlexoftService, ICodiempresa Codiempresa)
        {
            _context = context;
            _reservasService = reservasService;
            _UsuarioFlexoftService = UsuarioFlexoftService;
            _Codiempresa = Codiempresa;
        }

        [HttpGet("componentes-cierre/{idSubOr}")]
        public async Task<ActionResult<List<FormulaComponenteDTO>>> GetComponentesCierrePorReserva(string idSubOr)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(idSubOr))
                {
                    return BadRequest("El identificador de la orden de fabricación es requerido.");
                }

                var orden = await _context.Ordefabrs
                    .AsNoTracking()
                    .FirstOrDefaultAsync(x => x.IdSubOr == idSubOr);

                if (orden == null)
                {
                    return NotFound($"No se encontró la orden de fabricación {idSubOr}");
                }

                var canProy = orden.CanProy ?? 0m;

                var reservas = await _context.Reservas
                    .AsNoTracking()
                    .Where(r => r.IdSubOr == idSubOr)
                    .ToListAsync();

                if (reservas.Count == 0)
                {
                    return Ok(new List<FormulaComponenteDTO>());
                }

                var codigos = reservas
                    .Select(r => r.Cod_Inte)
                    .Distinct()
                    .ToList();

                var stocks = await _context.Movistos
                    .AsNoTracking()
                    .Where(m => codigos.Contains(m.Cod_Inte ?? 0))
                    .GroupBy(m => m.Cod_Inte ?? 0)
                    .Select(g => new
                    {
                        CodInte = g.Key,
                        Stock = g.Sum(x => (x.CantIngre ?? 0m) - (x.CantSalid ?? 0m))
                    })
                    .ToDictionaryAsync(x => x.CodInte, x => x.Stock);

                var resultado = reservas
                    .GroupBy(r => new
                    {
                        CodInte = r.Cod_Inte,
                        CodExte = r.Cod_Exte ?? string.Empty,
                        Descrip = r.Descrip ?? string.Empty
                    })
                    .OrderBy(g => g.Key.CodExte)
                    .Select(g =>
                    {
                        var cantRes = g.Sum(x => x.CantRes ?? 0m);
                        var usoUnit = g.Sum(x => x.UsoUnit ?? 0m);
                        var uso = usoUnit;

                        if (uso <= 0m && canProy >= 0.05m)
                        {
                            uso = cantRes / canProy;
                        }

                        stocks.TryGetValue(g.Key.CodInte, out var stock);

                        return new FormulaComponenteDTO
                        {
                            CodIElem = g.Key.CodInte,
                            CodXElem = g.Key.CodExte,
                            Descripcion = g.Key.Descrip,
                            UsoBruto = uso,
                            Stock = stock,
                            EsRepetido = false,
                            NecesidadTotal = 0m,
                            CantidadAUtilizar = 0m
                        };
                    })
                    .ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener componentes de cierre desde reservas: {ex.Message}");
            }
        }

        [HttpGet("pendientes-por-producto/{codint:int}")]
        public async Task<ActionResult<List<StockReservaPendienteDTO>>> GetPendientesPorProducto(int codint)
        {
            try
            {
                var agrupado = _context.Reservas
                    .AsNoTracking()
                    .Where(r => r.Cod_Inte == codint
                                && (r.CantCons ?? 0m) < ((r.CantRes ?? 0m) - 0.005m))
                    .GroupBy(r => new
                    {
                        IdSubOr = r.IdSubOr ?? string.Empty,
                        IdenLote = r.IdenLote ?? string.Empty,
                        Fecha = r.FechRes.HasValue ? r.FechRes.Value.Date : DateTime.MinValue
                    })
                    .Select(g => new
                    {
                        g.Key.IdSubOr,
                        g.Key.IdenLote,
                        FechaReserva = g.Key.Fecha,
                        CantRes = g.Sum(x => x.CantRes ?? 0m),
                        CantCons = g.Sum(x => x.CantCons ?? 0m)
                    });

                var query = from g in agrupado
                            join ofa in _context.Ordefabrs.AsNoTracking()
                                on g.IdSubOr equals (ofa.IdSubOr ?? string.Empty) into ofJoin
                            from ofa in ofJoin.DefaultIfEmpty()
                            select new StockReservaPendienteDTO
                            {
                                IdSubOr = g.IdSubOr,
                                IdenLote = g.IdenLote,
                                FechaReserva = g.FechaReserva == DateTime.MinValue ? null : g.FechaReserva,
                                FechaEntregaSolicitada = ofa != null ? ofa.FechEnSol : null,
                                CodExteDestino = ofa != null ? (ofa.Cod_Exte ?? string.Empty) : string.Empty,
                                DescripDestino = ofa != null ? (ofa.Descrip ?? string.Empty) : string.Empty,
                                CantReservada = g.CantRes,
                                CantConsumida = g.CantCons,
                                CantPendiente = (g.CantRes - g.CantCons) < 0m ? 0m : (g.CantRes - g.CantCons)
                            };

                var result = await query
                    .Where(x => x.CantPendiente > 0m)
                    .OrderBy(x => x.IdSubOr)
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener reservas pendientes: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza las reservas consumidas para una orden de fabricación específica
        /// Traducción de la rutina ACTUALIZARESERVA de VB6
        /// </summary>
        /// <param name="norfa">Número de orden de fabricación</param>
        /// <param name="codigoInterno">Código interno del producto</param>
        /// <param name="cantidad">Cantidad a consumir (positiva)</param>
        /// <returns>True si se actualizó correctamente</returns>
        [HttpPost("actualizar-consumo")]
        public async Task<ActionResult<bool>> ActualizarReserva(string norfa, int codigoInterno, decimal cantidad)
        {
            try
            {
                if (string.IsNullOrEmpty(norfa) || codigoInterno <= 0 || cantidad <= 0)
                {
                    return BadRequest("Parámetros inválidos para actualizar reserva");
                }

                // SALCON# = (-1) * CANTI - En VB6 se usa negativo, aquí usamos positivo
                decimal cantidadPorConsumir = cantidad;

                // Obtener todas las reservas de la orden y código interno específico
                var reservas = await _context.Reservas
                    .Where(r => r.IdSubOr == norfa && r.Cod_Inte == codigoInterno)
                    .ToListAsync();

                if (!reservas.Any())
                {
                    return NotFound($"No se encontraron reservas para la orden {norfa} y código {codigoInterno}");
                }

                // Procesar cada reserva hasta consumir toda la cantidad
                foreach (var reservaItem in reservas)
                {
                    if (cantidadPorConsumir <= 0)
                        break;

                    decimal cantidadReservada = reservaItem.CantRes ?? 0;
                    decimal cantidadConsumidaAnterior = reservaItem.CantCons ?? 0;
                    decimal disponibleParaConsumir = cantidadReservada - cantidadConsumidaAnterior;

                    if (disponibleParaConsumir > 0)
                    {
                        if (disponibleParaConsumir >= cantidadPorConsumir)
                        {
                            // La reserva actual puede cubrir toda la cantidad restante
                            reservaItem.CantCons = cantidadConsumidaAnterior + cantidadPorConsumir;
                            cantidadPorConsumir = 0;
                        }
                        else
                        {
                            // La reserva actual se consume completamente
                            cantidadPorConsumir -= disponibleParaConsumir;
                            reservaItem.CantCons = cantidadReservada;
                        }
                    }
                }

                // Guardar cambios
                await _context.SaveChangesAsync();

                return Ok(true);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al actualizar reserva: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene todas las reservas asociadas a una orden de fabricación
        /// </summary>
        /// <param name="idSubOr">Identificador de la orden de fabricación</param>
        /// <returns>Lista de reservas de la orden de fabricación</returns>
        [HttpGet("por-of/{idSubOr}")]
        public async Task<ActionResult> GetReservasPorOf(string idSubOr)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(idSubOr))
                {
                    return BadRequest("El identificador de la orden de fabricación es requerido.");
                }

                var reservas = await _reservasService.ObtenerReservasPorOfAsync(idSubOr);
                
                // Si no hay reservas, devolver lista vacía en lugar de NotFound
                // Esto es un comportamiento normal, no un error
                if (reservas == null)
                {
                    reservas = new List<Reserva>();
                }

                return Ok(reservas);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener las reservas: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza múltiples reservas de una orden de fabricación
        /// </summary>
        /// <param name="reservas">Lista de reservas con los datos actualizados</param>
        /// <returns>Resultado de la operación de actualización</returns>
        [HttpPut("multiples")]
        public async Task<ActionResult> ActualizarMultiplesReservas(List<Reserva> reservas)
        {
            try
            {
                int? idFlexoft = _UsuarioFlexoftService.InterObtenerIdFlexoft(User);
                short? codigoEmpresa = _Codiempresa.IObtenerCodiempresa(User);
                if (!idFlexoft.HasValue)
                {
                    return BadRequest("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                if (reservas == null || !reservas.Any())
                {
                    return BadRequest("La lista de reservas es requerida y no puede estar vacía.");
                }

                var resultados = new List<object>();
                int reservasActualizadas = 0;
                int reservasConError = 0;

                foreach (var reserva in reservas)
                {
                    try
                    {
                        if (reserva.Reserva_Id <= 0)
                        {
                            resultados.Add(new { reservaId = reserva.Reserva_Id, error = "ID de reserva inválido" });
                            reservasConError++;
                            continue;
                        }

                        var resultado = await _reservasService.ActualizarReservaAsync(reserva);
                        
                        if (resultado)
                        {
                            resultados.Add(new { reservaId = reserva.Reserva_Id, exito = true });
                            reservasActualizadas++;
                        }
                        else
                        {
                            resultados.Add(new { reservaId = reserva.Reserva_Id, error = "No se encontró la reserva" });
                            reservasConError++;
                        }
                    }
                    catch (Exception ex)
                    {
                        resultados.Add(new { reservaId = reserva.Reserva_Id, error = ex.Message });
                        reservasConError++;
                    }
                }

                return Ok(new 
                { 
                    mensaje = $"Proceso completado: {reservasActualizadas} reservas actualizadas, {reservasConError} con errores.",
                    reservasActualizadas,
                    reservasConError,
                    detalles = resultados
                });
            }
            catch (Exception ex)
            {
                return BadRequest($"Error general al actualizar las reservas: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza una reserva existente
        /// </summary>
        /// <param name="reserva">Objeto reserva con los datos actualizados</param>
        /// <returns>Resultado de la operación de actualización</returns>
        [HttpPut]
        public async Task<ActionResult> ActualizarReserva(Reserva reserva)
        {
            try
            {
                int? idFlexoft = _UsuarioFlexoftService.InterObtenerIdFlexoft(User);
                short? codigoEmpresa = _Codiempresa.IObtenerCodiempresa(User);
                if (!idFlexoft.HasValue)
                {
                    return BadRequest("(Sesion Caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                if (reserva == null)
                {
                    return BadRequest("Los datos de la reserva son requeridos.");
                }

                if (reserva.Reserva_Id <= 0)
                {
                    return BadRequest("El ID de la reserva debe ser válido.");
                }

                var resultado = await _reservasService.ActualizarReservaAsync(reserva);
                
                if (!resultado)
                {
                    return NotFound($"No se encontró la reserva con ID {reserva.Reserva_Id}.");
                }

                return Ok(new { mensaje = "Reserva actualizada exitosamente.", reserva });
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al actualizar la reserva: {ex.Message}");
            }
        }
    }
}
