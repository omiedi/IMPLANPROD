using System;

namespace IMPLANPROD.Shared.DTOs.Notificaciones
{
    public class NotificacionReglaSqlItemDTO
    {
        public int Id { get; set; }

        public string Titulo { get; set; } = string.Empty;

        public string? Resumen { get; set; }

        public bool Activa { get; set; }

        public int MaxFilasDetalle { get; set; }

        public int IntervaloMinutos { get; set; }

        public DateTime? UltimaEjecucion { get; set; }

        public string? UltimoResultadoHash { get; set; }

        public string? SqlConnectionName { get; set; }

        /// <summary>
        /// Modo de seguimiento: 0 = Hash (cambios en resultado), 1 = Incremental (nuevos registros)
        /// </summary>
        public int ModoSeguimiento { get; set; }

        /// <summary>
        /// Campo incremental para seguimiento de nuevos registros (ej: NroPed, IdOrden, etc.)
        /// Solo se usa cuando ModoSeguimiento = 1
        /// </summary>
        public string? CampoSeguimiento { get; set; }

        /// <summary>
        /// Último valor procesado en modo incremental
        /// </summary>
        public string? UltimoValorProcesado { get; set; }

        /// <summary>
        /// Si es true, la notificación se envía a todos los usuarios
        /// Si es false, se envía solo a los usuarios en DestinatariosIds
        /// </summary>
        public bool EsParaTodos { get; set; }

        /// <summary>
        /// Lista de IDs de usuarios destinatarios (separados por coma) cuando EsParaTodos = false
        /// </summary>
        public string? DestinatariosIds { get; set; }
    }
}
