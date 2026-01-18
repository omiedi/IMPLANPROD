using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.DTOs
{


    /// <summary>
    /// DTO para representar el stock de un repuesto en un depósito específico
    /// </summary>
    public class StockDepositoDTO
    {
        /// <summary>
        /// Identificador del depósito
        /// </summary>
        public int? Deposito { get; set; }

        /// <summary>
        /// Cantidad actual de stock en el depósito
        /// </summary>
        public decimal Stock { get; set; }

        public string? Descripcion { get; set; }
    }

    /// <summary>
    /// DTO para realizar ajustes de stock (ingresos o salidas)
    /// </summary>
    public class StockAjusteDTO
    {
        /// <summary>
        /// Identificador del repuesto
        /// </summary>
        [Required(ErrorMessage = "El identificador del repuesto es obligatorio")]
        public int? CodInte { get; set; }

        /// <summary>
        /// Código externo del repuesto
        /// </summary>
        [Required(ErrorMessage = "El código externo es obligatorio")]
        public string? CodExte { get; set; }

        /// <summary>
        /// Descripción del repuesto
        /// </summary>
        [Required(ErrorMessage = "La descripción es obligatoria")]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Identificador del depósito donde se realiza el ajuste
        /// </summary>
        [Required(ErrorMessage = "El depósito es obligatorio")]
        public int? Deposito { get; set; }

        /// <summary>
        /// Cantidad a ingresar (debe ser positiva o cero)
        /// </summary>
        [Range(0, double.MaxValue, ErrorMessage = "El ingreso debe ser un valor positivo")]
        public decimal Ingreso { get; set; }

        /// <summary>
        /// Cantidad a retirar (debe ser positiva o cero)
        /// </summary>
        [Range(0, double.MaxValue, ErrorMessage = "La salida debe ser un valor positivo")]
        public decimal Salida { get; set; }

        /// <summary>
        /// Referencia o motivo del ajuste
        /// </summary>
        [Required(ErrorMessage = "La referencia es obligatoria")]
        public string? Referencia { get; set; }

        public int? NumUsuar { get; set; }
        public short? CodiEmpr { get; set; }
        public string? IdenLote { get; set; }

    }
}

