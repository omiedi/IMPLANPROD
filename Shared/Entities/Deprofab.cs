using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa el detalle de un programa de fabricación.
    /// Contiene los productos y cantidades asociadas a cada programa.
    /// Mapea a la tabla DEPROFAB en la base de datos.
    /// Implementa IAuditableEntity para auditoría automática.
    /// </summary>
    [Table("DEPROFAB")]
    public class Deprofab : IAuditableEntity
    {
        /// <summary>
        /// Identificador único del detalle (campo IDENTITY)
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Número de programa al que pertenece este detalle (FK a PROFABRI)
        /// </summary>
        [Required]
        public int NumeroPrograma { get; set; }

        /// <summary>
        /// Código interno del producto (FK a MASTER.Codint)
        /// </summary>
        [Required]
        public int CodigoProducto { get; set; }

        /// <summary>
        /// Descripción del producto (copiado de MASTER para histórico)
        /// </summary>
        [MaxLength(64)]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Cantidad del producto en el programa
        /// </summary>
        [Required]
        [Column(TypeName = "decimal(18,4)")]
        public decimal Cantidad { get; set; }

        /// <summary>
        /// Observaciones específicas del item en el programa
        /// </summary>
        [MaxLength(52)]
        public string? Observaciones { get; set; }

        // Campos de auditoría - IAuditableEntity
        [Display(Name = "Fecha de Creación")]
        public DateTime? AddRecord { get; set; }

        [Display(Name = "Última Actualización")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa para auditoría .NET
        /// </summary>
        public int? CodiEmprNet { get; set; } = 0;

        /// <summary>
        /// Número de usuario que realizó la última modificación
        /// </summary>
        public int? NumUsuar { get; set; }
    }
}

/*
-- Script SQL para crear la tabla DEPROFAB

CREATE TABLE DEPROFAB (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    NumeroPrograma INT NOT NULL,
    CodigoProducto INT NOT NULL,
    Descripcion NVARCHAR(64) NULL,
    Cantidad DECIMAL(18,4) NOT NULL,
    Observaciones NVARCHAR(52) NULL,
    -- Campos de auditoría (se llenan automáticamente)
    AddRecord DATETIME NULL,
    LastUpdate DATETIME NULL,
    CodiEmprNet INT NULL DEFAULT 0,
    NumUsuar INT NULL,
    
    -- Foreign key a PROFABRI
    CONSTRAINT FK_DEPROFAB_PROFABRI 
        FOREIGN KEY (NumeroPrograma) 
        REFERENCES PROFABRI(NumeroPrograma)
        ON DELETE CASCADE
);

-- Índice para búsqueda por programa
CREATE INDEX IX_DEPROFAB_NumeroPrograma ON DEPROFAB(NumeroPrograma);

-- Índice para búsqueda por producto
CREATE INDEX IX_DEPROFAB_CodigoProducto ON DEPROFAB(CodigoProducto);

-- Índice compuesto para búsquedas combinadas
CREATE INDEX IX_DEPROFAB_Programa_Producto ON DEPROFAB(NumeroPrograma, CodigoProducto);
*/
