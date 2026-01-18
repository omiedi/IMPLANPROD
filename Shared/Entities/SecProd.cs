using DocumentFormat.OpenXml.Office2019.Excel.RichData;
using System.ComponentModel.DataAnnotations;



namespace IMPLANPROD.Shared.Entities
{
    public class SecProd
    {   [Key]
        public int Id { get; set; }

        [Display(Name = "Código SecProd")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public int CodigoSecProd { get; set; }

        [Display(Name = "Descripción")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(32, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? Descripcion { get; set; }

        [Display(Name = "Estado")]
        public int Status { get; set; } = 0;

        [Display(Name = "Deposito")]
        public int? Deposito { get; set; } = 0;

    }
}
