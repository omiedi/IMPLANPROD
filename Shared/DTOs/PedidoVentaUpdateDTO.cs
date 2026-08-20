using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para editar una Orden de Venta completa: encabezado (PEDENCA) + detalle (PEDDETA).
    /// Incluye el NroPed a editar.
    /// </summary>
    public class PedidoVentaUpdateDTO
    {
        /// <summary>
        /// Número de pedido a editar (dato de negocio)
        /// </summary>
        [Required]
        public int NroPed { get; set; }

        // -------------------- ENCABEZADO (PEDENCA) --------------------
        public DateTime? FeSolEnt { get; set; }
        [Required]
        public int NroClie { get; set; }
        [Required, StringLength(64)]
        public string RasoClie { get; set; } = string.Empty;
        [StringLength(32)]
        public string? NroOcom { get; set; }
        [Required]
        public short CondPag { get; set; }
        [Required]
        public short Vended { get; set; }
        public decimal? PorComVend { get; set; }
        public short ListPre { get; set; } = 1;
        /// <summary>
        /// Texto de descuento aplicado (se graba en PEDENCA.TDescApli)
        /// </summary>
        [StringLength(14)]
        public string? TDescApli { get; set; }
        [StringLength(64)]
        public string? Motivo { get; set; }
        public short? MoneEmis { get; set; }
        public string? Observa { get; set; }
        public decimal? TipoCam { get; set; }

        public int? NumUsuar { get; set; }
        public short? NTranspo { get; set; }
        [StringLength(48)]
        public string? TranspEnt { get; set; }



        // -------------------- DETALLE (PEDDETA) --------------------
        [Required]
        public List<PedidoVentaItemDTO> Items { get; set; } = new();
    }
}
