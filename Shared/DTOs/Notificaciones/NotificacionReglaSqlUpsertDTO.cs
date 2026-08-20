using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs.Notificaciones
{
    public class NotificacionReglaSqlUpsertDTO
    {
        public int Id { get; set; }

        [Required]
        [MaxLength(200)]
        public string Titulo { get; set; } = string.Empty;

        [MaxLength(500)]
        public string? Resumen { get; set; }

        public bool Activa { get; set; } = true;

        [Required]
        public string SqlQuery { get; set; } = string.Empty;

        public string? SqlConnectionName { get; set; }

        public int MaxFilasDetalle { get; set; } = 200;

        public int IntervaloMinutos { get; set; } = 15;

        /// <summary>
        /// Modo de seguimiento: 0 = Hash (cambios en resultado), 1 = Incremental (nuevos registros)
        /// </summary>
        public int ModoSeguimiento { get; set; } = 0;

        /// <summary>
        /// Campo incremental para seguimiento de nuevos registros (ej: NroPed, IdOrden, etc.)
        /// Solo se usa cuando ModoSeguimiento = 1
        /// </summary>
        [MaxLength(100)]
        public string? CampoSeguimiento { get; set; }

        /// <summary>
        /// Si es true, la notificación se envía a todos los usuarios
        /// Si es false, se envía solo a los usuarios en DestinatariosIds
        /// </summary>
        public bool EsParaTodos { get; set; } = true;

        /// <summary>
        /// Lista de IDs de usuarios destinatarios (separados por coma) cuando EsParaTodos = false
        /// Ej: "1,5,23,45"
        /// </summary>
        [MaxLength(2000)]
        public string? DestinatariosIds { get; set; }
    }
}
