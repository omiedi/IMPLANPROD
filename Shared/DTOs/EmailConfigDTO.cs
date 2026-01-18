using System.ComponentModel.DataAnnotations;


namespace IMPLANPROD.Shared.DTOs
{
    public class EmailConfigDTO
    {
        [Required(ErrorMessage = "El correo electrónico es obligatorio")]
        [EmailAddress(ErrorMessage = "El formato del correo electrónico no es válido")]
        public string EmailAddress { get; set; } = string.Empty;

        [Required(ErrorMessage = "La contraseña de aplicación es obligatoria")]
        [MinLength(16, ErrorMessage = "La contraseña de aplicación debe tener al menos 16 caracteres")]
        public string AppPassword { get; set; } = string.Empty;

        [Required(ErrorMessage = "El nombre para mostrar es obligatorio")]
        public string DisplayName { get; set; } = string.Empty;
    }
}


