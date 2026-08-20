using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    public class MasterAltaXCopiaDTO
    {
        [Required]
        public int BaseCodint { get; set; }

        [Required]
        [MaxLength(24)]
        public string NuevoCodigo { get; set; } = string.Empty;

        [Required]
        [MaxLength(64)]
        public string NuevaDescripcion { get; set; } = string.Empty;

        [MaxLength(4)]
        public string? NuevaUnidadMedida { get; set; }

        public short? NuevoTipoMaterial { get; set; }
    }
}
