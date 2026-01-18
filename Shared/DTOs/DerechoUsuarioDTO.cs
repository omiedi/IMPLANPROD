namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar los derechos de un usuario sobre una aplicación
    /// Replica la lógica del código VB6 para mostrar derechos
    /// </summary>
    public class DerechoUsuarioDTO
    {
        /// <summary>
        /// Código de la aplicación
        /// </summary>
        public string CodApli { get; set; } = string.Empty;

        /// <summary>
        /// Descripción de la aplicación
        /// </summary>
        public string DesApli { get; set; } = string.Empty;

        /// <summary>
        /// Número de usuario (puede ser null si no tiene derechos asignados)
        /// </summary>
        public short? NumUser { get; set; }

        /// <summary>
        /// Derecho de lectura: 1 = Con derecho, 0 = Sin derecho
        /// </summary>
        public short DerRead { get; set; } = 0;

        /// <summary>
        /// Derecho de escritura: 1 = Con derecho, 0 = Sin derecho
        /// </summary>
        public short DerWrite { get; set; } = 0;

        /// <summary>
        /// Indica si tiene derecho de lectura (para binding con checkbox)
        /// </summary>
        public bool TieneLectura
        {
            get => DerRead == 1;
            set => DerRead = (short)(value ? 1 : 0);
        }

        /// <summary>
        /// Indica si tiene derecho de escritura (para binding con checkbox)
        /// </summary>
        public bool TieneEscritura
        {
            get => DerWrite == 1;
            set => DerWrite = (short)(value ? 1 : 0);
        }

        /// <summary>
        /// ID del registro DER_ACCE (si existe)
        /// </summary>
        public int? IdDerAcce { get; set; }
    }
}
