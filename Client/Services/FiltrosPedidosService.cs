using System;

namespace IMPLANPROD.Client.Services
{
    /// <summary>
    /// Servicio para persistir los filtros de pedidos durante la navegación
    /// Mantiene el estado de los filtros cuando el usuario navega entre páginas
    /// </summary>
    public class FiltrosPedidosService
    {
        /// <summary>
        /// Filtro de texto general (N° OC, Referencia, etc.)
        /// </summary>
        public string Filter { get; set; } = string.Empty;

        /// <summary>
        /// Cliente seleccionado para filtrar
        /// </summary>
        public string ClienteSeleccionado { get; set; } = string.Empty;

        /// <summary>
        /// ID del cliente seleccionado
        /// </summary>
        public int? ClienteSeleccionadoId { get; set; }

        /// <summary>
        /// Estado del pedido para filtrar
        /// </summary>
        public string FiltroEstado { get; set; } = string.Empty;

        /// <summary>
        /// Fecha desde para filtrar
        /// </summary>
        public DateTime? FechaDesde { get; set; }

        /// <summary>
        /// Fecha hasta para filtrar
        /// </summary>
        public DateTime? FechaHasta { get; set; }

        /// <summary>
        /// Página actual en la paginación
        /// </summary>
        public int PaginaActual { get; set; } = 1;

        /// <summary>
        /// Limpia todos los filtros
        /// </summary>
        public void LimpiarFiltros()
        {
            Filter = string.Empty;
            ClienteSeleccionado = string.Empty;
            ClienteSeleccionadoId = null;
            FiltroEstado = string.Empty;
            FechaDesde = null;
            FechaHasta = null;
            PaginaActual = 1;
        }

        /// <summary>
        /// Guarda el estado actual de los filtros
        /// </summary>
        public void GuardarEstado(string filter, string clienteSeleccionado, int? clienteSeleccionadoId, 
            string filtroEstado, DateTime? fechaDesde, DateTime? fechaHasta, int paginaActual)
        {
            Filter = filter;
            ClienteSeleccionado = clienteSeleccionado;
            ClienteSeleccionadoId = clienteSeleccionadoId;
            FiltroEstado = filtroEstado;
            FechaDesde = fechaDesde;
            FechaHasta = fechaHasta;
            PaginaActual = paginaActual;
        }

        /// <summary>
        /// Indica si hay filtros activos
        /// </summary>
        public bool TieneFiltrosActivos()
        {
            return !string.IsNullOrEmpty(Filter) ||
                   !string.IsNullOrEmpty(ClienteSeleccionado) ||
                   !string.IsNullOrEmpty(FiltroEstado) ||
                   FechaDesde.HasValue ||
                   FechaHasta.HasValue;
        }
    }
}
