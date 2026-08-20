using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para transferencia múltiple de stock entre depósitos
    /// </summary>
    public class StockTransferenciaMultipleDTO
    {
        [Required(ErrorMessage = "El depósito origen es requerido")]
        [Range(1, int.MaxValue, ErrorMessage = "Seleccione un depósito origen válido")]
        public int DepositoOrigen { get; set; }

        [Required(ErrorMessage = "El depósito destino es requerido")]
        [Range(1, int.MaxValue, ErrorMessage = "Seleccione un depósito destino válido")]
        public int DepositoDestino { get; set; }

        [Required(ErrorMessage = "La fecha es requerida")]
        public DateTime Fecha { get; set; }

        [Required(ErrorMessage = "La referencia es obligatoria")]
        [MaxLength(32, ErrorMessage = "La referencia no puede superar los 32 caracteres")]
        public string Referencia { get; set; } = string.Empty;

        [Required(ErrorMessage = "Debe agregar al menos un producto")]
        [MinLength(1, ErrorMessage = "Debe agregar al menos un producto")]
        public List<StockTransferenciaMultipleItemDTO> Items { get; set; } = new();
    }

    /// <summary>
    /// DTO para cada ítem de la transferencia múltiple
    /// </summary>
    public class StockTransferenciaMultipleItemDTO
    {
        [Required]
        public int CodigoInterno { get; set; }

        public string? CodigoExterno { get; set; }

        public string? Descripcion { get; set; }

        [Required(ErrorMessage = "La cantidad es requerida")]
        [Range(0.01, double.MaxValue, ErrorMessage = "La cantidad debe ser mayor a 0")]
        public decimal Cantidad { get; set; }

        [MaxLength(20, ErrorMessage = "El lote no puede superar los 20 caracteres")]
        public string? Lote { get; set; }

        [MaxLength(20, ErrorMessage = "La OC no puede superar los 20 caracteres")]
        public string? OrdenCompra { get; set; }
    }
}
