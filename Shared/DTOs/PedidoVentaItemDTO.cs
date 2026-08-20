using System;
using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para item/detalle de pedido de venta.
    /// Contiene los campos mínimos necesarios para grabar en PEDDETA.
    /// </summary>
    public class PedidoVentaItemDTO
    {
        /// <summary>
        /// Orden del item dentro del pedido (1..N)
        /// </summary>
        public short NuOrItem { get; set; }

        /// <summary>
        /// Código interno del producto (MASTER.CODINT)
        /// </summary>
        public int? CodiInt { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        [StringLength(24)]
        public string? CodiExt { get; set; }

        /// <summary>
        /// Sub código (si aplica)
        /// </summary>
        [StringLength(8)]
        public string? SubCod { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del producto
        /// </summary>
        [StringLength(384)]
        public string? Descrip { get; set; }

        /// <summary>
        /// Color, talle y dimensiones (si aplican)
        /// </summary>
        [StringLength(16)] public string? TColor { get; set; } = string.Empty;
        [StringLength(16)] public string? TTalle { get; set; } = string.Empty;
        public decimal? TLargo { get; set; } = 0;
        public decimal? TAncho { get; set; } = 0;
        public decimal? TEspesor { get; set; } = 0;

        /// <summary>
        /// Moneda y tipo de cambio (si el item necesita diferenciarlo)
        /// </summary>
        public short? MoneEmis { get; set; }
        public decimal? TipoCam { get; set; }

        /// <summary>
        /// Lista, precio, cantidades e importes
        /// </summary>
        public short? ListaDeItem { get; set; } = 1;
        public decimal? PreList { get; set; }
        public decimal? CanAsig { get; set; }
        public decimal? CanPed { get; set; }
        public decimal? PreUnid { get; set; }
        public decimal? ImpoTot { get; set; }

        /// <summary>
        /// Vendedor representante (1er carácter de MOTIVO: 'V'=Venta, 'R'=Reparación, etc.)
        /// </summary>
        [StringLength(8)]
        public string? VenRep { get; set; }

        /// <summary>
        /// Número de serie y observaciones
        /// </summary>
        [StringLength(32)] public string? NumeroSerie { get; set; } = string.Empty;
        [StringLength(512)] public string? ObservaItem { get; set; } = string.Empty;

        /// <summary>
        /// Señalización de sin cargo (0/1)
        /// </summary>
        public short? SinCargo { get; set; } = 0;

        /// <summary>
        /// Fecha solicitada de entrega del item
        /// </summary>
        public DateTime? FeSolEnt { get; set; }

        /// <summary>
        /// TPEDITEM: Identificador del item (formato: NRO_PEDIDO-NRO_ITEM)
        /// Ejemplo: "12345-1", "12345-2", etc.
        /// </summary>
        [StringLength(32)]
        public string? TPEDITEM { get; set; }
    }
}
