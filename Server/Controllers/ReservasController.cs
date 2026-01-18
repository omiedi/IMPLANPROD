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
