using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para transferencia de stock con transformación de código:
    /// sale una cantidad del producto origen (Cantsalid) e ingresa la misma cantidad
    /// del producto destino (Cantingr), posiblemente en depósitos distintos.
    /// </summary>
    public class StockTransferenciaTransformacionDTO
    {
        [Required(ErrorMessage = "El código de origen es requerido")]
        public int? CodigoOrigenInt { get; set; }

        public string? CodigoOrigenExt { get; set; }

        public string? DescripcionOrigen { get; set; }

        [Required(ErrorMessage = "El depósito de salida es requerido")]
        [Range(1, int.MaxValue, ErrorMessage = "Seleccione un depósito de salida válido")]
        public int DepositoSalida { get; set; }

        [MaxLength(20, ErrorMessage = "El lote no puede superar los 20 caracteres")]
        public string? Lote { get; set; }

        [Required(ErrorMessage = "El código de destino es requerido")]
        public int? CodigoDestinoInt { get; set; }

        public string? CodigoDestinoExt { get; set; }

        public string? DescripcionDestino { get; set; }

        [Required(ErrorMessage = "El depósito de ingreso es requerido")]
        [Range(1, int.MaxValue, ErrorMessage = "Seleccione un depósito de ingreso válido")]
        public int DepositoIngreso { get; set; }

        [Required(ErrorMessage = "La fecha es requerida")]
        public DateTime Fecha { get; set; } = DateTime.Today;

        [Required(ErrorMessage = "La cantidad es requerida")]
        [Range(0.01, double.MaxValue, ErrorMessage = "La cantidad debe ser mayor a 0")]
        public decimal Cantidad { get; set; }

        [Required(ErrorMessage = "La referencia es requerida")]
        [MaxLength(32, ErrorMessage = "La referencia no puede superar los 32 caracteres")]
        public string? Referencia { get; set; }

        /// <summary>
        /// Stock disponible en el depósito de salida para el código origen (solo informativo).
        /// </summary>
        public decimal StockOrigen { get; set; }
    }
}
