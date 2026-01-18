using System;

namespace Shared.Entities
{
    /// <summary>
    /// Entidad para almacenar la definición y configuración de reportes personalizados.
    /// Permite guardar la estructura del reporte, condiciones y formato de columnas en formato JSON.
    /// </summary>
    public class ReportePersonalizado
    {
        /// <summary>
        /// Identificador único del reporte.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Nombre descriptivo del reporte, definido por el usuario.
        /// </summary>
        public string Nombre { get; set; }

        /// <summary>
        /// Nombre de usuario que creó el reporte (opcional).
        /// </summary>
        public string? Usuario { get; set; }

        /// <summary>
        /// Configuración completa del reporte en formato JSON (tablas, campos, relaciones, formato de columnas, etc).
        /// </summary>
        public string ConfiguracionJson { get; set; }

        /// <summary>
        /// Fecha de creación del reporte.
        /// </summary>
        public DateTime FechaCreacion { get; set; }

        /// <summary>
        /// Fecha de la última modificación del reporte.
        /// </summary>
        public DateTime FechaModificacion { get; set; }
    }
}
