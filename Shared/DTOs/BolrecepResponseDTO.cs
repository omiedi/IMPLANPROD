using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO que representa la respuesta del backend después de crear una boleta de recepción.
    /// Contiene los datos generados que son necesarios para registrar los movimientos de stock.
    /// </summary>
    public class BolrecepResponseDTO
    {
        /// <summary>
        /// Número de boleta de recepción generado por el sistema
        /// </summary>
        public int Numebore { get; set; }

        /// <summary>
        /// Identificador único del lote generado para este material
        /// </summary>
        public string Idenlote { get; set; }

        /// <summary>
        /// Código externo del material recibido
        /// </summary>
        public string Cod_Exte { get; set; }

        /// <summary>
        /// Cantidad recibida del material
        /// </summary>
        public decimal CantReci { get; set; }
    }
}
