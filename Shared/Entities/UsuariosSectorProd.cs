using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Security.Principal;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa la relación muchos a muchos entre Usuario y SecProd
    /// Permite asignar múltiples sectores productivos a un usuario
    /// </summary>
    public class UsuariosSectorProd : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador único de la relación
        /// </summary>
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Identificador del usuario en el sistema Flexoft
        /// </summary>
        [Display(Name = "ID Flexoft")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public int IdFlexoft { get; set; }

        /// <summary>
        /// Nombre del usuario
        /// </summary>
        [Display(Name = "Nombre Usuario")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(50, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? NombreUsuario { get; set; }

        /// <summary>
        /// Código del sector productivo
        /// </summary>
        [Display(Name = "Código SecProd")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public int CodigoSecProd { get; set; }

        /// <summary>
        /// Descripción del sector productivo
        /// </summary>
        [Display(Name = "Descripción")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(32, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Estado del registro (0 = activo, -1 = eliminado)
        /// </summary>
        [Display(Name = "Estado")]
        public int Status { get; set; } = 0;

        /// <summary>
        /// Fecha de creación del registro
        /// </summary>
        [Display(Name = "Fecha Creación")]
        public DateTime FechaCreacion { get; set; } = DateTime.Now;

        /// <summary>
        /// Navegación hacia la entidad Usuario
        /// </summary>
        public Usuario? Usuario { get; set; }

        /// <summary>
        /// Navegación hacia la entidad SecProd
        /// </summary>
        public SecProd? SecProd { get; set; }

        [Display(Name = "SecProdId ")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public int SecProdId { get; set; }


        [Display(Name = "UsuarioId ")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public int UsuarioId { get; set; }

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

//        CREATE TABLE UsuariosSectorProd(
//  Id INT IDENTITY(1,1) PRIMARY KEY,
//  IdFlexoft INT NOT NULL,
//    NombreUsuario NVARCHAR(50),
//    CodigoSecProd INT NOT NULL,
//    Descripcion NVARCHAR(32) NOT NULL,
//    Status INT NOT NULL DEFAULT 0,
//    FechaCreacion DATETIME NOT NULL DEFAULT GETDATE(),
//    SecProdId INT NOT NULL,
//    UsuarioId INT NOT NULL
//);
    }
}
