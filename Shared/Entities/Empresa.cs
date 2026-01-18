using IMPLANPROD.Shared.Entities;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    public class Empresa : ITimestampEntity, IAuditableEntity
    {
        public int Id { get; set; }

        [Display(Name = "Empresa")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(100)]
        public string razon { get; set; } = null!;


        [Display(Name = "Fantasia")]
        [MaxLength(100, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? fantasia { get; set; }

        [Display(Name = "CUIT")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? CUIT { get; set; }


        [Display(Name = "Calle")]
        [MaxLength(100, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? calle { get; set; }

        [Display(Name = "Número")]
        [MaxLength(100, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? numero { get; set; }

        [Display(Name = "Complemento")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? complemento { get; set; }

        [Display(Name = "Cód.Provincia")]
        [MaxLength(100, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public int? provinciaemp { get; set; }

        [Display(Name = "Localidad")]
        [MaxLength(70, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? localidad { get; set; }

        [Display(Name = "Cód.Postal")]
        [MaxLength(24, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? CP { get; set; }

        [Display(Name = "Télefono 1")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? TEL1 { get; set; }

        [Display(Name = "Télefono 21")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? TEL2 { get; set; }


        [Display(Name = "Fax")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? fax { get; set; }

        [Display(Name = "Mail")]
        [MaxLength(255, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? email { get; set; }

        [Display(Name = "Web")]
        [MaxLength(2048, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? WEB { get; set; }

        [Display(Name = "Notas")]
        [MaxLength(1073741, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? notasemp
        {
            get; set;

        }

        [Display(Name = "Estado")]
        public int? status { get; set; } = 0;
        public ICollection<Dependencia>? Dependencias { get; set; }
        public int DependenciasNumber => Dependencias == null ? 0 : Dependencias.Count; //para leer la cant de dependencias que tiene una empresa

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}


