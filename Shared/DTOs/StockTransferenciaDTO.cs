using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs

{
    public class StockTransferenciaDTO
    {
        public int CodigoRepuesto { get; set; }
        public string? CodExt { get; set; }
        public string? Descripcion { get; set; }

        [Required(ErrorMessage = "El depósito origen es requerido")]
        [Range(1, int.MaxValue, ErrorMessage = "Seleccione un depósito origen válido")]
        public int DepositoOrigen { get; set; }

        [Required(ErrorMessage = "El depósito destino es requerido")]
        [Range(1, int.MaxValue, ErrorMessage = "Seleccione un depósito destino válido")]
        public int DepositoDestino { get; set; }

        [Required(ErrorMessage = "La cantidad a transferir es requerida")]
        [Range(0.01, double.MaxValue, ErrorMessage = "La cantidad debe ser mayor a 0")]
        public decimal CantidadTransferir { get; set; }

        public decimal StockOrigen { get; set; }

        [MaxLength(32, ErrorMessage = "La referencia no puede superar los 32 caracteres")]
        public string? Referencia { get; set; }

        public int? NumUsuar { get; set; }
        public short? CodiEmpr { get; set; }


    }
}