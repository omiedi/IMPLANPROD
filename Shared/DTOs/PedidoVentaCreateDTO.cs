using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para crear una Orden de Venta completa: encabezado (PEDENCA) + detalle (PEDDETA).
    /// Solo incluye los campos necesarios según requerimiento.
    /// </summary>
    public class PedidoVentaCreateDTO
    {
        // -------------------- ENCABEZADO (PEDENCA) --------------------

        /// <summary>
        /// Fecha de entrega solicitada
        /// </summary>
        public DateTime? FeSolEnt { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        [Required]
        public int NroClie { get; set; }

        /// <summary>
        /// Razón social del cliente
        /// </summary>
        [Required, StringLength(64)]
        public string RasoClie { get; set; } = string.Empty;

        /// <summary>
        /// Número de orden de compra
        /// </summary>
        [StringLength(32)]
        public string? NroOcom { get; set; }

        /// <summary>
        /// Condición de pago (código)
        /// </summary>
        [Required]
        public short CondPag { get; set; }

        /// <summary>
        /// Vendedor (código)
        /// </summary>
        [Required]
        public short Vended { get; set; }

        /// <summary>
        /// Porcentaje de comisión del vendedor
        /// </summary>
        public decimal? PorComVend { get; set; }

        /// <summary>
        /// Lista de precios a usar
        /// </summary>
        public short ListPre { get; set; } = 1;

        /// <summary>
        /// Texto de descuento aplicado (se graba en PEDENCA.TDescApli)
        /// Sugerencia: almacenar el porcentaje con 4 decimales en formato texto.
        /// </summary>
        [StringLength(14)]
        public string? TDescApli { get; set; }

        /// <summary>
        /// Motivo (primer carácter se usará como VenRep en ítems)
        /// </summary>
        [StringLength(64)]
        public string? Motivo { get; set; }

        /// <summary>
        /// Moneda de emisión (si aplica al encabezado)
        /// </summary>
        public short? MoneEmis { get; set; }

        /// <summary>
        /// Observaciones del pedido (PEDENCA.Observa)
        /// </summary>
        public string? Observa { get; set; }

        /// <summary>
        /// Tipo de cambio (si aplica al encabezado)
        /// </summary>
        public decimal? TipoCam { get; set; }
        /// <summary>
        /// nro transporte (si aplica al encabezado)
        /// </summary>
        public short? NTranspo { get; set; }

        /// <summary>
        /// Descripcion transporte (si aplica al encabezado)
        /// </summary>
        [StringLength(48)]
        public string? TranspEnt { get; set; }

        /// <summary>
        /// Domicilio de entrega (PEDENCA.DomiEnt)
        /// </summary>
        [StringLength(128)]
        public string? DomiEnt { get; set; }

        /// <summary>
        /// Localidad de entrega (PEDENCA.LocaEnt)
        /// </summary>
        [StringLength(64)]
        public string? LocaEnt { get; set; }

        /// <summary>
        /// nro usuario
        /// </summary>
        public int? NumUsuar { get; set; }


        // -------------------- DETALLE (PEDDETA) --------------------

        /// <summary>
        /// Ítems del detalle de la orden de venta
        /// </summary>
        [Required]
        public List<PedidoVentaItemDTO> Items { get; set; } = new();
    }
}
