using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el procesamiento de cierre de orden de fabricación con control de calidad
    /// Contiene todos los parámetros necesarios para ejecutar MOVISTOCK_APROF
    /// </summary>
    public class CierreCalidadDTO
    {
        /// <summary>
        /// Modo de cierre de calidad (debe ser 1 para activar el procesamiento)
        /// </summary>
        [Required(ErrorMessage = "El modo de cierre de calidad es requerido")]
        public int ModoCierreCalidad { get; set; }

        /// <summary>
        /// Fecha del movimiento de stock
        /// </summary>
        [Required(ErrorMessage = "La fecha es requerida")]
        public DateTime Fecha { get; set; }

        /// <summary>
        /// Código interno del producto (equivalente a CICONJ% en VB6)
        /// </summary>
        [Required(ErrorMessage = "El código interno es requerido")]
        public int CodigoInterno { get; set; }

        /// <summary>
        /// Identificador del lote (equivalente a LOTE$ en VB6)
        /// </summary>
        [Required(ErrorMessage = "El lote es requerido")]
        [MaxLength(50, ErrorMessage = "El lote no puede exceder 50 caracteres")]
        public string Lote { get; set; } = string.Empty;

        /// <summary>
        /// Código del tipo de movimiento (equivalente a CMOV$ en VB6)
        /// </summary>
        [Required(ErrorMessage = "El código de movimiento es requerido")]
        [MaxLength(10, ErrorMessage = "El código de movimiento no puede exceder 10 caracteres")]
        public string CodigoMovimiento { get; set; } = string.Empty;

        /// <summary>
        /// Código primario (equivalente a DOPRI$ en VB6)
        /// </summary>
        [Required(ErrorMessage = "El código primario es requerido")]
        [MaxLength(50, ErrorMessage = "El código primario no puede exceder 50 caracteres")]
        public string Dopri { get; set; } = string.Empty;

        /// <summary>
        /// Código secundario (equivalente a DOSEC$ en VB6)
        /// </summary>
        [Required(ErrorMessage = "El código secundario es requerido")]
        [MaxLength(50, ErrorMessage = "El código secundario no puede exceder 50 caracteres")]
        public string Dosec { get; set; } = string.Empty;

        /// <summary>
        /// Referencia del movimiento (equivalente a REFE$ en VB6)
        /// </summary>
        [MaxLength(100, ErrorMessage = "La referencia no puede exceder 100 caracteres")]
        public string Referencia { get; set; } = string.Empty;

        /// <summary>
        /// Valor unitario del producto (equivalente a VUNI# en VB6)
        /// </summary>
        [Required(ErrorMessage = "El valor unitario es requerido")]
        [Range(0, double.MaxValue, ErrorMessage = "El valor unitario debe ser mayor o igual a 0")]
        public decimal ValorUnitario { get; set; }

        /// <summary>
        /// Número de cliente (equivalente a NC% en VB6)
        /// Puede ser negativo para proveedores
        /// </summary>
        public int NumeroCliente { get; set; }

        /// <summary>
        /// Depósito de origen (equivalente a DEPX% en VB6)
        /// </summary>
        [Required(ErrorMessage = "El depósito es requerido")]
        public int Deposito { get; set; }

        /// <summary>
        /// Cantidad a mover (equivalente a CANTI en VB6)
        /// </summary>
        [Required(ErrorMessage = "La cantidad es requerida")]
        public decimal Cantidad { get; set; }

        /// <summary>
        /// Depósito de consumo (equivalente a DP2% en VB6)
        /// </summary>
        public int DepositoConsumo { get; set; }

        /// <summary>
        /// Identificador de texto opcional (equivalente a CLA0$ en VB6)
        /// Si está vacío, se generará automáticamente
        /// </summary>
        [MaxLength(50, ErrorMessage = "El ID de texto no puede exceder 50 caracteres")]
        public string IdTexto { get; set; } = string.Empty;
    }
}
