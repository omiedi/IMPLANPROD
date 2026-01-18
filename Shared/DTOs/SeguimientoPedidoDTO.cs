namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el seguimiento completo de un pedido (órdenes de fabricación y remitos)
    /// </summary>
    public class SeguimientoPedidoDTO
    {
        public int NroPed { get; set; }
        public List<OrdenFabricacionSeguimientoDTO> OrdenesFabricacion { get; set; } = new();
        public List<RemitoSeguimientoDTO> Remitos { get; set; } = new();
    }

    /// <summary>
    /// DTO para información de órdenes de fabricación relacionadas al pedido
    /// </summary>
    public class OrdenFabricacionSeguimientoDTO
    {
        /// <summary>
        /// Código externo de la orden de fabricación
        /// </summary>
        public string? CodExte { get; set; }

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string? Descrip { get; set; }

        /// <summary>
        /// Cantidad proyectada
        /// </summary>
        public decimal? CanProy { get; set; }

        /// <summary>
        /// Cantidad aprobada
        /// </summary>
        public decimal? CanApro { get; set; }

        /// <summary>
        /// Identificación del lote
        /// </summary>
        public string? IdenLote { get; set; }

        /// <summary>
        /// Estado de la orden de fabricación
        /// </summary>
        public short? StatuOrf { get; set; }

        /// <summary>
        /// Descripción del estado
        /// </summary>
        public string EstadoDescripcion
        {
            get
            {
                return StatuOrf switch
                {
                    0 => "Activa",
                    1 => "Cerrada",
                    2 => "Anulada",
                    _ => "Desconocido"
                };
            }
        }

        /// <summary>
        /// Porcentaje de avance (CanApro / CanProy * 100)
        /// </summary>
        public decimal PorcentajeAvance
        {
            get
            {
                if (CanProy.HasValue && CanProy.Value > 0 && CanApro.HasValue)
                {
                    return Math.Round((CanApro.Value / CanProy.Value) * 100, 2);
                }
                return 0;
            }
        }
    }

    /// <summary>
    /// DTO para información de remitos relacionados al pedido
    /// </summary>
    public class RemitoSeguimientoDTO
    {
        /// <summary>
        /// Número de remito (DOPRILAR)
        /// </summary>
        public string? NumeroRemito { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        public string? CodExte { get; set; }

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string? Descrip { get; set; }

        /// <summary>
        /// Cantidad de salida
        /// </summary>
        public decimal? CantSalid { get; set; }

        /// <summary>
        /// Fecha del movimiento
        /// </summary>
        public DateTime? FechMov { get; set; }
    }
}
