using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para enviar los parámetros de vista previa de reportes dinámicos.
    /// </summary>
    public class VistaPreviaDTO
    {
        /// <summary>
        /// Nombre de la tabla principal.
        /// </summary>
        public string TablaPrincipal { get; set; }

        /// <summary>
        /// Lista de campos seleccionados de la tabla principal.
        /// </summary>
        public List<string> Campos { get; set; } = new List<string>();

        /// <summary>
        /// Lista de filtros aplicados.
        /// </summary>
        public List<FiltroDTO> Filtros { get; set; } = new List<FiltroDTO>();

        /// <summary>
        /// Lista de joins a aplicar en la consulta.
        /// </summary>
        public List<JoinDTO> Joins { get; set; } = new List<JoinDTO>();
    }

    /// <summary>
    /// DTO para definir un join entre tablas en la vista previa.
    /// </summary>
    public class JoinDTO
    {
        /// <summary>
        /// Nombre de la tabla a unir.
        /// </summary>
        public string Tabla { get; set; }

        /// <summary>
        /// Campo de la tabla principal para la unión.
        /// </summary>
        public string CampoPrincipal { get; set; }

        /// <summary>
        /// Campo de la tabla unida (foránea) para la unión.
        /// </summary>
        public string CampoForanea { get; set; }

        /// <summary>
        /// Tipo de join: LEFT, RIGHT, INNER
        /// </summary>
        public string TipoJoin { get; set; } = "INNER";

        /// <summary>
        /// Lista de campos de la tabla foránea seleccionados para mostrar en la vista previa.
        /// </summary>
        public List<string> CamposSeleccionados { get; set; } = new List<string>();
    }

    /// <summary>
    /// DTO para definir un filtro en la vista previa.
    /// </summary>
    public class FiltroDTO
    {
        /// <summary>
        /// Nombre del campo a filtrar.
        /// </summary>
        public string Campo { get; set; }
        /// <summary>
        /// Operador (=, >, <, LIKE, etc).
        /// </summary>
        public string Operador { get; set; }
        /// <summary>
        /// Valor del filtro.
        /// </summary>
        public object Valor { get; set; }
    }
}
