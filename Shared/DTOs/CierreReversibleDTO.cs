namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar un cierre reversible de orden de fabricación
    /// 
    /// Este DTO se usa para mostrar al usuario los cierres que pueden ser anulados.
    /// Un cierre es reversible cuando tiene CantIngre > 0 en la tabla MOVISTO con CodiMovi = 'OF'.
    /// 
    /// Flujo de uso:
    /// 1. El usuario hace clic en "Anulación x Cierre" en la tabla de órdenes
    /// 2. Se llama al endpoint GET api/ordefabr/cierres-reversibles/{codInte}
    /// 3. El endpoint retorna una lista de estos DTOs
    /// 4. El usuario selecciona uno de la grilla
    /// 5. Se llama al endpoint PUT api/ordefabr/anular-cierre/{idMovimiento} para anularlo
    /// </summary>
    public class CierreReversibleDTO
    {
        /// <summary>
        /// ID del movimiento en la tabla MOVISTO
        /// Este ID se usa para identificar el cierre específico a anular
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Documento primario corto (IdSubOr de la orden de fabricación)
        /// Es el identificador de la orden de fabricación, ejemplo: "OF-000123005"
        /// </summary>
        public string DoPriCor { get; set; } = string.Empty;

        /// <summary>
        /// Fecha del movimiento de stock
        /// Es la fecha en que se realizó el cierre de la orden
        /// </summary>
        public DateTime? FechMov { get; set; }

        /// <summary>
        /// Cantidad ingresada en el cierre
        /// Esta es la cantidad que se va a restar de la cantidad aprobada al anular el cierre
        /// </summary>
        public decimal CantIngre { get; set; }

        /// <summary>
        /// Fecha y hora de registro del movimiento
        /// Indica cuándo se creó el registro en la base de datos
        /// </summary>
        public DateTime? FeReMovi { get; set; }

        /// <summary>
        /// Lote del producto fabricado
        /// Identifica el lote específico que se cerró en esta operación
        /// </summary>
        public string? IdenLote { get; set; }
    }
}
