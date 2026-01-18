using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    public class Dependencia : ITimestampEntity, IAuditableEntity
    {
        public int Id { get; set; }


        [Display(Name = "Nombre Filial")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(70, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? dependencia { get; set; }

        [Display(Name = "Calle")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? calledep { get; set; }


        [Display(Name = "Número")]
        [MaxLength(10, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? numerodep { get; set; }

        [Display(Name = "Complemento")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? complementodep { get; set; }

        [Display(Name = "Cód.Provincia")]
        public int? provinciadep { get; set; }

        [Display(Name = "Localidad")]
        [MaxLength(70, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? localidaddep { get; set; }

        [Display(Name = "Cód.Postal")]
        [MaxLength(24, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? CPdep { get; set; }

        [Display(Name = "Télefono 1")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? TEL1dep { get; set; }

        [Display(Name = "Télefono 2")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? TEL2dep { get; set; }


        [Display(Name = "Fax")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? faxdep { get; set; }

        [Display(Name = "Mail")]
        [MaxLength(255, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? emaildep { get; set; }


        [Display(Name = "Notas")]
        [MaxLength(1073741, ErrorMessage = "El campo {0} no puede tener  más de {1} caracteres.")]
        public string? notasdep { get; set; }

        // Clave foránea para la relación con Empresa
        public int EmpresaId { get; set; } //agregado lluego de que se cargaron las relaciones 

        public Empresa? Empresa { get; set; } // Navegación hacia Empresa

        // Relación uno a muchos con Sectores
        public ICollection<Sector>? Sectores { get; set; }
        public int SectoresNumber => Sectores == null ? 0 : Sectores.Count; //para leer la cant de sectores que tiene una dependencia

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
