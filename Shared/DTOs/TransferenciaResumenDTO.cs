using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Resumen de una transferencia para la pantalla de Re-Impresiones.
    /// </summary>
    public class TransferenciaResumenDTO
    {
        /// <summary>
        /// Número completo de la transferencia (DoPriCor), ej. TR.61113
        /// </summary>
        public string Numero { get; set; } = string.Empty;

        /// <summary>
        /// Código secundario (DoSecCor), ej. TR.260825
        /// </summary>
        public string DoSecCor { get; set; } = string.Empty;

        /// <summary>
        /// Fecha de la transferencia
        /// </summary>
        public DateTime Fecha { get; set; }

        /// <summary>
        /// Indica si la transferencia contiene más de un producto
        /// </summary>
        public bool EsMultiple { get; set; }

        /// <summary>
        /// Código externo del producto (solo cuando es un ítem único)
        /// </summary>
        public string? CodigoExterno { get; set; }

        /// <summary>
        /// Descripción del producto (solo cuando es un ítem único)
        /// </summary>
        public string? Descripcion { get; set; }

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
        /// Nombre del usuario que realizó la transferencia
        /// </summary>
        public string UsuarioNombre { get; set; } = string.Empty;
    }
}
