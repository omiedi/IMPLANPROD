using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad Secoper. Representa la tabla [SECOPER] con todos los campos especificados.
    /// Todos los campos float se mapean como decimal para mayor precisión.
    /// </summary>
    public class Secoper : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        public int CodIConj { get; set; }
        public short? CodiEmpr { get; set; }

        [StringLength(24)]
        public string? CodXConj { get; set; }

        public short? NumeOper { get; set; }
        public short? AltOper { get; set; }

        [StringLength(64)]
        public string? DescOper { get; set; }

        public short? StatOper { get; set; }

        [StringLength(6)]
        public string? CodMaq0 { get; set; }

        [StringLength(6)]
        public string? CodMaq1 { get; set; }

        [StringLength(6)]
        public string? CodMaq2 { get; set; }

        [StringLength(6)]
        public string? CodMaq3 { get; set; }

        [StringLength(6)]
        public string? CodMaq4 { get; set; }

        [StringLength(6)]
        public string? CodMaq5 { get; set; }

        // Campos convertidos de float a decimal(18,4)
        [Column(TypeName = "decimal(18,4)")]
        public decimal? HorsPrep { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? HorsFijo { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? MinuStan { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? HorsMHer { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? HorsLimp { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? PieCiclo { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? PorDemor { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Cadencia { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? CoefSolp { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? LoteStan { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantiOps { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Cantiops2 { get; set; }

        public short? CategOps { get; set; }
        public short? Categops2 { get; set; }

        [StringLength(24)]
        public string? ProgCNum { get; set; }

        [StringLength(24)]
        public string? HoProFab { get; set; }

        // Campo ntext mapeado como string
        [StringLength(512)]
        public string? ObseOper { get; set; }

        public short? ProTerOp { get; set; }

        [StringLength(12)]
        public string? ProvOper { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? PreUnid { get; set; }

        public short? MonePrec { get; set; }
        public DateTime? FechPrec { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Nuoper_Prep { get; set; }
        public short? Catoper_Prep { get; set; }
        public DateTime? Fech_Revis { get; set; }

        [StringLength(32)]
        public string? Usua_Revis { get; set; }

        [Required]
        public Guid rowguid { get; set; }

        public short? SecuenciaFabricacion { get; set; }
        public short? OperacionTerceros { get; set; }
        public short? NuOrdenManual { get; set; }

        /// <summary>
        /// Indica si la operación es nueva y aún no ha sido guardada en la base de datos
        /// </summary>
        [NotMapped]
        public bool EsNueva { get; set; }

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
//ALTER TABLE dbo.SECOPER ADD Id INT IDENTITY(1,1) PRIMARY KEY;

//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN HorsPrep DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN HorsFijo DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN MinuStan DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN HorsMHer DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN HorsLimp DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN PieCiclo DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN PorDemor DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN Cadencia DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN CoefSolp DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN LoteStan DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN CantiOps DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN Cantiops2 DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN PreUnid DECIMAL(18,4) NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN Nuoper_Prep DECIMAL(18,4) NULL;

//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN ObseOper NVARCHAR(512) NULL;

//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN FechPrec DATETIME NULL;
//ALTER TABLE dbo.SECOPER 
//ALTER COLUMN Fech_Revis DATETIME NULL;
//ALTER TABLE SECOPER
//ALTER COLUMN CodIConj INT;