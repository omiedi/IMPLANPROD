using System;
using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el Vale de Transferencias de Stock que se muestra en vista previa e impresión.
    /// Origen: resultados de transferencias simples, múltiples o por pedido.
    /// </summary>
    public class ValeTransferenciaDTO
    {
        /// <summary>
        /// Número del vale, equivalente a DoPriCor (ej. TR.61113)
        /// </summary>
        public string Numero { get; set; } = string.Empty;

        /// <summary>
        /// Código secundario con formato TR.YYMMDD
        /// </summary>
        public string DoSecCor { get; set; } = string.Empty;

        /// <summary>
        /// Fecha de la transferencia
        /// </summary>
        public DateTime Fecha { get; set; }

        /// <summary>
        /// Identificador del depósito origen
        /// </summary>
        public int DepositoOrigen { get; set; }

        /// <summary>
        /// Nombre del depósito origen
        /// </summary>
        public string DepositoOrigenNombre { get; set; } = string.Empty;

        /// <summary>
        /// Identificador del depósito destino
        /// </summary>
        public int DepositoDestino { get; set; }

        /// <summary>
        /// Nombre del depósito destino
        /// </summary>
        public string DepositoDestinoNombre { get; set; } = string.Empty;

        /// <summary>
        /// Motivo o referencia de la transferencia
        /// </summary>
        public string Motivo { get; set; } = string.Empty;

        /// <summary>
        /// Ítems transferidos
        /// </summary>
        public List<ValeTransferenciaItemDTO> Items { get; set; } = new();
    }

    /// <summary>
    /// Ítem del Vale de Transferencias de Stock.
    /// </summary>
    public class ValeTransferenciaItemDTO
    {
        /// <summary>
        /// Código externo del producto
        /// </summary>
        public string CodigoExterno { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>
        /// Unidad de medida del producto (master.Umedida)
        /// </summary>
        public string UnidadMedida { get; set; } = string.Empty;

        /// <summary>
        /// Número de pedido, si aplica
        /// </summary>
        public int? NroPed { get; set; }

        /// <summary>
        /// Cantidad transferida
        /// </summary>
        public decimal Cantidad { get; set; }

        /// <summary>
        /// Lote, si aplica
        /// </summary>
        public string? Lote { get; set; }
    }
}
