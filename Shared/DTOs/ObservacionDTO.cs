using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar una observación genérica de cualquier tabla
    /// </summary>
    public class ObservacionDTO
    {
        /// <summary>
        /// Identificador único de la observación
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Fecha de creación de la observación
        /// </summary>
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Contenido de la observación
        /// </summary>
        public string? Observaciones { get; set; }

        /// <summary>
        /// Usuario que creó la observación
        /// </summary>
        public int? NumUsuar { get; set; }
    }
}
