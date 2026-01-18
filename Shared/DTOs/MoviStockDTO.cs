using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el registro de movimientos de stock que replica la funcionalidad de MOVISTOK en VB6
    /// Incluye todos los parámetros necesarios para el proceso de movimiento de stock
    /// </summary>
    public class MoviStockDTO
    {
        /// <summary>
        /// Fecha del movimiento (FECHA% en VB6)
        /// </summary>
        [Required(ErrorMessage = "La fecha del movimiento es obligatoria")]
        public DateTime Fecha { get; set; }

        /// <summary>
        /// Código interno del producto (CI% en VB6)
        /// </summary>
        [Required(ErrorMessage = "El código interno del producto es obligatorio")]
        public int CodigoInterno { get; set; }

        /// <summary>
        /// Lote del producto (LOTE$ en VB6)
        /// </summary>
        [StringLength(16)]
        public string? Lote { get; set; } = string.Empty;

        /// <summary>
        /// Código de movimiento (CMOV$ en VB6)
        /// </summary>
        [Required(ErrorMessage = "El código de movimiento es obligatorio")]
        [StringLength(2)]
        public string CodigoMovimiento { get; set; } = string.Empty;

        /// <summary>
        /// Documento primario largo (DOPRILA$ en VB6)
        /// </summary>
        [StringLength(24)]
        public string? DocumentoPrimarioLargo { get; set; } = string.Empty;

        /// <summary>
        /// Documento primario corto (DOPRICO$ en VB6)
        /// </summary>
        [StringLength(12)]
        public string? DocumentoPrimarioCorto { get; set; } = string.Empty;

        /// <summary>
        /// Documento secundario largo (DOSELA$ en VB6)
        /// </summary>
        [StringLength(24)]
        public string? DocumentoSecundarioLargo { get; set; } = string.Empty;

        /// <summary>
        /// Documento secundario corto (DOSECO$ en VB6)
        /// </summary>
        [StringLength(12)]
        public string? DocumentoSecundarioCorto { get; set; } = string.Empty;

        /// <summary>
        /// Referencia del movimiento (REFE$ en VB6)
        /// </summary>
        [StringLength(64)]
        public string? Referencia { get; set; } = string.Empty;

        /// <summary>
        /// Número de orden item (NUORIT% en VB6)
        /// </summary>
        public int? NumeroOrdenItem { get; set; } = 0;

        /// <summary>
        /// Valor unitario (VUNI# en VB6)
        /// </summary>
        public decimal? ValorUnitario { get; set; } = 0;

        /// <summary>
        /// Moneda (MONEII% en VB6)
        /// </summary>
        public short? Moneda { get; set; } = 0;

        /// <summary>
        /// Número de cliente (NC% en VB6) - Si es negativo se trata como proveedor
        /// </summary>
        public int? NumeroCliente { get; set; } = 0;

        /// <summary>
        /// Depósito (DEPO% en VB6)
        /// </summary>
        [Required(ErrorMessage = "El depósito es obligatorio")]
        public int Deposito { get; set; }

        /// <summary>
        /// Cantidad del movimiento (CANTI en VB6) - Positivo=Ingreso, Negativo=Salida
        /// </summary>
        [Required(ErrorMessage = "La cantidad es obligatoria")]
        public decimal Cantidad { get; set; }

        /// <summary>
        /// Número de pedido cliente (NUPECLIEN& en VB6) - Parámetro opcional
        /// </summary>
        public long? NumeroPedidoCliente { get; set; }

        /// <summary>
        /// Número orden pedido (NORITPE% en VB6) - Parámetro opcional
        /// </summary>
        public short? NumeroOrdenPedido { get; set; }

        /// <summary>
        /// Descripción larga (DELARGA$ en VB6) - Parámetro opcional
        /// </summary>
        public string? DescripcionLarga { get; set; }

        /// <summary>
        /// Documento de compra (DOCOMPRA$ en VB6) - Parámetro opcional
        /// </summary>
        [StringLength(24)]
        public string? DocumentoCompra { get; set; }

        /// <summary>
        /// Identificador de movimiento (IDMOVIM$ en VB6) - Parámetro opcional
        /// </summary>
        [StringLength(32)]
        public string? IdentificadorMovimiento { get; set; }

        /// <summary>
        /// Precio con IVA (PRECONIVA# en VB6) - Parámetro opcional
        /// </summary>
        public decimal? PrecioConIva { get; set; }

        /// <summary>
        /// Tipo de transferencia (TIPOTRAN% en VB6) - Parámetro opcional
        /// </summary>
        public short? TipoTransferencia { get; set; }

        /// <summary>
        /// Orden de compra retorno abierto (ORDCOM_RTOABIER$ en VB6) - Parámetro opcional
        /// </summary>
        [StringLength(64)]
        public string? OrdenCompraRetornoAbierto { get; set; }

        /// <summary>
        /// ID de calidad (IdCalidad& en VB6) - Parámetro opcional
        /// </summary>
        public long? IdCalidad { get; set; }

        /// <summary>
        /// ID de operación (IDOPERACION% en VB6) - Parámetro opcional
        /// </summary>
        public short? IdOperacion { get; set; }

        /// <summary>
        /// Metros (METROS# en VB6) - Parámetro opcional
        /// </summary>
        public decimal? Metros { get; set; }

        /// <summary>
        /// Peso en kilogramos (PESOKG# en VB6) - Parámetro opcional
        /// </summary>
        public decimal? PesoKg { get; set; }

        /// <summary>
        /// Número de remito (DOPRINUM en MOVISTO) - Últimos 8 dígitos del remito
        /// Campo numérico que almacena el número secuencial del remito
        /// </summary>
        public int? NumeroRemito { get; set; }

        /// <summary>
        /// Prefijo del remito (NUMPVTA en MOVISTO) - Primeros 4 dígitos del remito
        /// Campo numérico que almacena el punto de venta o prefijo
        /// </summary>
        public short? PrefijoRemito { get; set; }
    }
}
