using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para manejar las operaciones relacionadas con las reservas de órdenes de fabricación
    /// </summary>
    public class ReservasService
    {
        private readonly DataContext _context;
        private readonly ILogger<ReservasService> _logger;

        /// <summary>
        /// Constructor del servicio de reservas
        /// </summary>
        /// <param name="context">Contexto de base de datos</param>
        /// <param name="logger">Logger para registrar eventos y errores</param>
        public ReservasService(DataContext context, ILogger<ReservasService> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene todas las reservas asociadas a una orden de fabricación específica
        /// </summary>
        /// <param name="idSubOr">Identificador de la orden de fabricación</param>
        /// <returns>Lista de reservas de la orden de fabricación</returns>
        public async Task<List<Reserva>> ObtenerReservasPorOfAsync(string idSubOr)
        {
            try
            {
                // Obtener todas las reservas que pertenecen a la orden de fabricación especificada
                var reservas = await _context.Reservas
                    .Where(r => r.IdSubOr == idSubOr)
                    .OrderBy(r => r.NuOrItem) // Ordenar por número de item
                    .ToListAsync();

                return reservas;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener las reservas de la OF {IdSubOr}", idSubOr);
                throw new Exception($"Error al obtener las reservas de la OF {idSubOr}: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Actualiza una reserva existente con lógica compleja de cálculo de consumos y saldos
        /// </summary>
        /// <param name="reserva">Objeto reserva con los datos actualizados</param>
        /// <returns>True si la actualización fue exitosa, False en caso contrario</returns>
        public async Task<bool> ActualizarReservaAsync(Reserva reserva)
        {
            try
            {
                // Verificar que la reserva existe
                var reservaExistente = await _context.Reservas
                    .FirstOrDefaultAsync(r => r.Reserva_Id == reserva.Reserva_Id);

                if (reservaExistente == null)
                {
                    _logger.LogWarning("No se encontró la reserva con ID {ReservaId}", reserva.Reserva_Id);
                    return false; // La reserva no existe
                }

                // Obtener datos de la orden de fabricación desde ORDEFABR
                var ordenFabricacion = await _context.Ordefabrs
                    .FirstOrDefaultAsync(of => of.IdSubOr == reserva.IdSubOr);

                if (ordenFabricacion == null)
                {
                    _logger.LogWarning("No se encontró la orden de fabricación {IdSubOr}", reserva.IdSubOr);
                    throw new Exception($"No se encontró la orden de fabricación {reserva.IdSubOr}");
                }

                // Obtener valores de la orden de fabricación
                var codInte = ordenFabricacion.Cod_Inte ?? 0;
                var canProy = ordenFabricacion.CanProy ?? 0;
                var canApro = ordenFabricacion.CanApro ?? 0;
                
                // Calcular capacidad pendiente (CAPEN)
                var caPen = canProy - canApro;
                if (caPen < 0) caPen = 0;

                _logger.LogInformation("Datos OF {IdSubOr}: Cod_Inte={CodInte}, CanProy={CanProy}, CanApro={CanApro}, CAPEN={CaPen}", 
                    reserva.IdSubOr, codInte, canProy, canApro, caPen);

                // Verificar si hay códigos repetidos en las reservas de esta OF
                var codigosRepetidos = await _context.Reservas
                    .Where(r => r.IdSubOr == reserva.IdSubOr)
                    .GroupBy(r => r.Cod_Inte)
                    .Where(g => g.Count() > 1)
                    .Select(g => new { CodInte = g.Key, Repeticiones = g.Count() })
                    .ToDictionaryAsync(x => x.CodInte, x => x.Repeticiones);

                // Calcular el consumo total desde MOVISTO
                // Equivalente a: VALOBADA("MOVISTO", "SUM(CantSalid)", "Cod_Inte = " & CIIX% & " AND CodiMovi = 'CF' AND DoPriLar = '" & NORFA$ & "'", "NOMESS")
                var consumoTotal = await _context.Movistos
                    .Where(m => m.Cod_Inte == reserva.Cod_Inte && 
                               m.CodiMovi == "CF" && 
                               m.DoPriLar == reserva.IdSubOr)
                    .SumAsync(m => m.CantSalid ?? 0);

                _logger.LogInformation("Consumo total calculado para Cod_Inte {CodInte} en OF {IdSubOr}: {ConsumoTotal}", 
                    reserva.Cod_Inte, reserva.IdSubOr, consumoTotal);

                // Si hay códigos repetidos, repartir el consumo
                decimal consumoFinal = consumoTotal;
                if (codigosRepetidos.ContainsKey(reserva.Cod_Inte))
                {
                    var repeticiones = codigosRepetidos[reserva.Cod_Inte];
                    consumoFinal = consumoTotal / repeticiones;
                    
                    _logger.LogInformation("Código repetido {CodInte} tiene {Repeticiones} repeticiones. Consumo repartido: {ConsumoFinal}", 
                        reserva.Cod_Inte, repeticiones, consumoFinal);
                }

                // Actualizar los campos básicos de la reserva
                reservaExistente.Cod_Inte = reserva.Cod_Inte;
                reservaExistente.Cod_Exte = reserva.Cod_Exte;
                reservaExistente.Descrip = reserva.Descrip;
                reservaExistente.IdSubOr = reserva.IdSubOr;
                reservaExistente.NuOrItem = reserva.NuOrItem;
                reservaExistente.Status = reserva.Status;
                reservaExistente.FechRes = DateTime.Now; // Equivalente a Now en VB6
                
                // Aplicar la nueva cantidad reservada (CANTX# en VB6)
                var cantidadReservada = reserva.CantRes ?? 0;
                reservaExistente.CantRes = cantidadReservada;
                
                // Asignar el consumo calculado
                reservaExistente.CantCons = consumoFinal;
                
                // Calcular el saldo: CANTX# - CANCONSUM
                var saldoEntre = cantidadReservada - consumoFinal;
                if (saldoEntre < 0) saldoEntre = 0;
                reservaExistente.Saldo_Entre = saldoEntre;
                
                _logger.LogInformation("Cálculo final - CantRes: {CantRes}, CantCons: {CantCons}, Saldo_Entre: {SaldoEntre}", 
                    cantidadReservada, consumoFinal, saldoEntre);

                // Actualizar otros campos según el patrón VB6
                reservaExistente.IdenLote = " "; // Equivalente a " " en VB6
                reservaExistente.Deposito = 0;
                
                // Calcular UsoUnit si CanProy > 0
                if (canProy >= 0.05m)
                {
                    reservaExistente.UsoUnit = cantidadReservada / canProy;
                }
                else
                {
                    reservaExistente.UsoUnit = 0;
                }
                
                reservaExistente.Nume_Solint = 0;
                
                // Actualizar campos opcionales si vienen en la reserva
                if (!string.IsNullOrEmpty(reserva.Observa))
                    reservaExistente.Observa = reserva.Observa;
                if (!string.IsNullOrEmpty(reserva.PNL_Des))
                    reservaExistente.PNL_Des = reserva.PNL_Des;
                if (!string.IsNullOrEmpty(reserva.PLC_Has))
                    reservaExistente.PLC_Has = reserva.PLC_Has;
                if (reserva.CodiEmpr.HasValue)
                    reservaExistente.CodiEmpr = reserva.CodiEmpr;
                if (reserva.Nume_Ocom.HasValue)
                    reservaExistente.Nume_Ocom = reserva.Nume_Ocom;
                if (reserva.Marborra.HasValue)
                    reservaExistente.Marborra = reserva.Marborra;
                if (!string.IsNullOrEmpty(reserva.NUMEROSERIE))
                    reservaExistente.NUMEROSERIE = reserva.NUMEROSERIE;
                if (!string.IsNullOrEmpty(reserva.UbicaLot))
                    reservaExistente.UbicaLot = reserva.UbicaLot;

                // Guardar los cambios
                _context.Reservas.Update(reservaExistente);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Reserva {ReservaId} actualizada exitosamente", reserva.Reserva_Id);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar la reserva con ID {ReservaId}", reserva.Reserva_Id);
                throw new Exception($"Error al actualizar la reserva con ID {reserva.Reserva_Id}: {ex.Message}", ex);
            }
        }
    }
}
