using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para manejar la distribución de IVA en órdenes de compra
    /// Equivalente a los arrays IMPOGRA e IMPOIVA del VB6
    /// </summary>
    public class DistribucionIvaDTO
    {
        /// <summary>
        /// Base imponible para 0% (Exento/No Gravado)
        /// Equivalente a IMPOGRA(0) del VB6
        /// </summary>
        public decimal BaseImponible0 { get; set; } = 0m;

        /// <summary>
        /// Base imponible para 10.5%
        /// Equivalente a IMPOGRA(1) del VB6
        /// </summary>
        public decimal BaseImponible105 { get; set; } = 0m;

        /// <summary>
        /// Base imponible para 21%
        /// Equivalente a IMPOGRA(2) del VB6
        /// </summary>
        public decimal BaseImponible21 { get; set; } = 0m;

        /// <summary>
        /// Base imponible para 27%
        /// Equivalente a IMPOGRA(3) del VB6
        /// </summary>
        public decimal BaseImponible27 { get; set; } = 0m;

        /// <summary>
        /// Importe de IVA calculado para 10.5%
        /// Equivalente a IMPOIVA(1) del VB6
        /// </summary>
        public decimal ImporteIva105 { get; set; } = 0m;

        /// <summary>
        /// Importe de IVA calculado para 21%
        /// Equivalente a IMPOIVA(2) del VB6
        /// </summary>
        public decimal ImporteIva21 { get; set; } = 0m;

        /// <summary>
        /// Importe de IVA calculado para 27%
        /// Equivalente a IMPOIVA(3) del VB6
        /// </summary>
        public decimal ImporteIva27 { get; set; } = 0m;

        /// <summary>
        /// Total de base imponible (suma de todas las bases)
        /// </summary>
        public decimal TotalBaseImponible => BaseImponible0 + BaseImponible105 + BaseImponible21 + BaseImponible27;

        /// <summary>
        /// Total de IVA (suma de todos los importes de IVA)
        /// </summary>
        public decimal TotalIva => ImporteIva105 + ImporteIva21 + ImporteIva27;

        /// <summary>
        /// Valida que la distribución sea correcta
        /// </summary>
        /// <param name="totalEsperado">Total que debe coincidir con la suma de bases imponibles</param>
        /// <returns>True si la distribución es válida</returns>
        public bool EsValida(decimal totalEsperado)
        {
            var diferencia = Math.Abs(TotalBaseImponible - totalEsperado);
            return diferencia <= 0.01m; // Tolerancia de 1 centavo
        }

        /// <summary>
        /// Obtiene los datos de IVA como array para compatibilidad con lógica VB6
        /// Índice 0: 0%, Índice 1: 10.5%, Índice 2: 21%, Índice 3: 27%
        /// </summary>
        /// <returns>Array con las bases imponibles</returns>
        public decimal[] ObtenerBasesImponibles()
        {
            return new decimal[] { BaseImponible0, BaseImponible105, BaseImponible21, BaseImponible27 };
        }

        /// <summary>
        /// Obtiene los importes de IVA como array para compatibilidad con lógica VB6
        /// Índice 0: 0%, Índice 1: 10.5%, Índice 2: 21%, Índice 3: 27%
        /// </summary>
        /// <returns>Array con los importes de IVA</returns>
        public decimal[] ObtenerImportesIva()
        {
            return new decimal[] { 0m, ImporteIva105, ImporteIva21, ImporteIva27 };
        }

        /// <summary>
        /// Obtiene las tasas de IVA correspondientes
        /// Índice 0: 0%, Índice 1: 10.5%, Índice 2: 21%, Índice 3: 27%
        /// </summary>
        /// <returns>Array con las tasas de IVA</returns>
        public decimal[] ObtenerTasasIva()
        {
            return new decimal[] { 0m, 0.105m, 0.21m, 0.27m };
        }
    }
}
