using IMPLANPROD.Shared.DTOs;

namespace IMPLANPROD.Shared.Interfaces
{
    /// <summary>
    /// Interfaz para el servicio de gestión de costos de items de órdenes de compra
    /// </summary>
    public interface IOCOMDETACostoService
    {
        /// <summary>
        /// Obtiene los costos de items de una orden de compra específica
        /// Combina datos de OCOMDETA con OCOMDETACosto
        /// Solo retorna items con StatOcom < 9
        /// </summary>
        /// <param name="numeOcom">Número de orden de compra</param>
        /// <returns>Lista de items con sus costos</returns>
        Task<List<OCOMDETACostoDto>> GetCostosPorOrdenAsync(int numeOcom);

        /// <summary>
        /// Guarda o actualiza el costo de un item
        /// Si existe el registro (match por NumeOcom + NuOrdItem + CodInte), actualiza
        /// Si no existe, crea uno nuevo
        /// </summary>
        /// <param name="dto">Datos del costo a guardar</param>
        /// <returns>DTO actualizado con el ID</returns>
        Task<OCOMDETACostoDto> GuardarCostoAsync(OCOMDETACostoDto dto);

        /// <summary>
        /// Guarda o actualiza múltiples costos en una sola operación
        /// Útil para guardar todos los items de una orden de una vez
        /// </summary>
        /// <param name="dtos">Lista de costos a guardar</param>
        /// <returns>Lista de DTOs actualizados</returns>
        Task<List<OCOMDETACostoDto>> GuardarCostosAsync(List<OCOMDETACostoDto> dtos);

        /// <summary>
        /// Guarda costos y registra cambios en NOVEDADES si la configuración ACOSREP lo permite
        /// Si ACOSREP es "SIEMPRE" o "PROHABIT", registra el cambio de precio en NOVEDADES
        /// Convierte precios a pesos usando la cotización actual
        /// </summary>
        /// <param name="dto">Datos de costos y observación</param>
        /// <returns>Lista de DTOs actualizados</returns>
        Task<List<OCOMDETACostoDto>> GuardarCostosConNovedadesAsync(GuardarCostosConNovedadesDto dto);
    }
}
