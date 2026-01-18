using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para manejar conversiones de moneda en pedidos de venta
    /// </summary>
    public class ConversionMonedaService
    {
        private readonly DataContext _context;
        private readonly ILogger<ConversionMonedaService> _logger;

        public ConversionMonedaService(DataContext context, ILogger<ConversionMonedaService> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene la cotización de una moneda específica
        /// </summary>
        /// <param name="codigoMoneda">Código de la moneda</param>
        /// <returns>Cotización de la moneda</returns>
        public async Task<decimal> ObtenerCotizacionAsync(int codigoMoneda)
        {
            try
            {
                var moneda = await _context.Tamoneds
                    .FirstOrDefaultAsync(m => m.CodigoMoneda == codigoMoneda && m.Status >= 0);

                if (moneda == null)
                {
                    _logger.LogWarning("No se encontró cotización para la moneda {CodigoMoneda}", codigoMoneda);
                    return 1; // Cotización por defecto
                }

                return moneda.Cotizacion;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener cotización para moneda {CodigoMoneda}", codigoMoneda);
                return 1; // Cotización por defecto en caso de error
            }
        }

        /// <summary>
        /// Obtiene información completa de una moneda
        /// </summary>
        /// <param name="codigoMoneda">Código de la moneda</param>
        /// <returns>Información de la moneda</returns>
        public async Task<Tamoned?> ObtenerMonedaAsync(int codigoMoneda)
        {
            try
            {
                return await _context.Tamoneds
                    .FirstOrDefaultAsync(m => m.CodigoMoneda == codigoMoneda && m.Status >= 0);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener información de moneda {CodigoMoneda}", codigoMoneda);
                return null;
            }
        }

        /// <summary>
        /// Convierte un valor de una moneda a otra
        /// </summary>
        /// <param name="valor">Valor a convertir</param>
        /// <param name="monedaOrigen">Código de moneda origen</param>
        /// <param name="monedaDestino">Código de moneda destino</param>
        /// <returns>Valor convertido</returns>
        public async Task<decimal> ConvertirValorAsync(decimal valor, int monedaOrigen, int monedaDestino)
        {
            try
            {
                // Si son la misma moneda, no hay conversión
                if (monedaOrigen == monedaDestino)
                    return valor;

                var cotizacionOrigen = await ObtenerCotizacionAsync(monedaOrigen);
                var cotizacionDestino = await ObtenerCotizacionAsync(monedaDestino);

                // Convertir a pesos (moneda base) y luego a la moneda destino
                var valorEnPesos = valor * cotizacionOrigen;
                var valorConvertido = valorEnPesos / cotizacionDestino;

                _logger.LogInformation("Conversión: {Valor} (moneda {MonedaOrigen}) = {ValorConvertido} (moneda {MonedaDestino})",
                    valor, monedaOrigen, valorConvertido, monedaDestino);

                return Math.Round(valorConvertido, 4);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al convertir valor {Valor} de moneda {MonedaOrigen} a {MonedaDestino}",
                    valor, monedaOrigen, monedaDestino);
                return valor; // Devolver valor original en caso de error
            }
        }

        /// <summary>
        /// Convierte todos los precios de un detalle de pedido a una nueva moneda
        /// </summary>
        /// <param name="detalles">Lista de detalles del pedido</param>
        /// <param name="monedaOrigen">Moneda actual de los precios</param>
        /// <param name="monedaDestino">Moneda a la que convertir</param>
        /// <returns>Lista de detalles con precios convertidos</returns>
        public async Task<List<Peddeta>> ConvertirDetallesPedidoAsync(List<Peddeta> detalles, int monedaOrigen, int monedaDestino)
        {
            try
            {
                if (monedaOrigen == monedaDestino)
                    return detalles; // No hay conversión necesaria

                var cotizacionOrigen = await ObtenerCotizacionAsync(monedaOrigen);
                var cotizacionDestino = await ObtenerCotizacionAsync(monedaDestino);

                foreach (var detalle in detalles)
                {
                    // Convertir PreList (precio de lista)
                    if (detalle.PreList.HasValue)
                    {
                        detalle.PreList = await ConvertirValorAsync(detalle.PreList.Value, monedaOrigen, monedaDestino);
                    }

                    // Convertir PreUnid (precio unitario)
                    if (detalle.PreUnid.HasValue)
                    {
                        detalle.PreUnid = await ConvertirValorAsync(detalle.PreUnid.Value, monedaOrigen, monedaDestino);
                    }

                    // Recalcular ImpoTot (importe total)
                    if (detalle.CanPed.HasValue && detalle.PreUnid.HasValue)
                    {
                        detalle.ImpoTot = Math.Round(detalle.CanPed.Value * detalle.PreUnid.Value, 2);
                    }

                    // Actualizar moneda del detalle
                    detalle.MoneEmis = (short)monedaDestino;
                }

                _logger.LogInformation("Convertidos {CantidadItems} items de moneda {MonedaOrigen} a {MonedaDestino}",
                    detalles.Count, monedaOrigen, monedaDestino);

                return detalles;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al convertir detalles de pedido de moneda {MonedaOrigen} a {MonedaDestino}",
                    monedaOrigen, monedaDestino);
                return detalles; // Devolver detalles originales en caso de error
            }
        }

        /// <summary>
        /// Obtiene todas las monedas activas del sistema
        /// </summary>
        /// <returns>Lista de monedas activas</returns>
        public async Task<List<Tamoned>> ObtenerMonedasActivasAsync()
        {
            try
            {
                return await _context.Tamoneds
                    .Where(m => m.Status >= 0)
                    .OrderBy(m => m.CodigoMoneda)
                    .ToListAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener monedas activas");
                return new List<Tamoned>();
            }
        }
    }
}
