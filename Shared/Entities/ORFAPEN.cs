using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("ORFAPEN")]
    public class Orfapen : ITimestampEntity, IAuditableEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }                   // Identity Primary Key
        public DateTime? FechMov { get; set; }
        public short? CodiEmpr { get; set; }
        public int? Cod_Inte { get; set; }
        public string? Cod_Exte { get; set; }
        public string? Descrip { get; set; }
        public string? IdenLote { get; set; }
        public string? CodiMovi { get; set; }
        public string? DoPriCor { get; set; }
        public string? DoPriLar { get; set; }
        public string? DoSecCor { get; set; }
        public string? DoSecLar { get; set; }
        public string? ReferCor { get; set; }
        public decimal? ValoUnit { get; set; }       // decimal(18,4)
        public short? MoneVal { get; set; }
        public int? DepoMovi { get; set; }            // int
        public short? DepoConsu { get; set; }
        public int? Nume_Clie { get; set; }
        public short? Nume_Prov { get; set; }
        public decimal? CantIngre { get; set; }       // decimal(18,4)
        public decimal? CantSalid { get; set; }       // decimal(18,4)
        public decimal? CantSaldo { get; set; }       // decimal(18,4)
        public DateTime? FeReMovi { get; set; }
        public int? NumUsuar { get; set; }
        public short? STATUS { get; set; }
        public string? IDTEXT { get; set; }
        public string? IDSITUACION { get; set; }

        [NotMapped] // Esta propiedad no se mapeará a la base de datos
        public decimal Ajuste { get; set; }

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

        // NOTA: NumUsuar ya existe en la línea 36, no se duplica aquí
    }
}


//ALTER TABLE[dbo].[ORFAPEN] ALTER COLUMN[ValoUnit] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORFAPEN] ALTER COLUMN[CantIngre] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORFAPEN] ALTER COLUMN[CantSalid] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORFAPEN] ALTER COLUMN[CantSaldo] decimal(18,4) NULL;

//ALTER TABLE[dbo].[ORFAPEN] ALTER COLUMN[DepoMovi] int NULL;

//ALTER TABLE[dbo].[ORFAPEN]
//ADD Id INT IDENTITY(1,1) NOT NULL;
//ALTER TABLE[dbo].[ORFAPEN]
//ADD CONSTRAINT PK_ORFAPEN_Id PRIMARY KEY (Id);

//MODIFICAR A INT EL NRO DE USUARIO
//ALTER TABLE ORFAPEN
//ADD NumUsuar  INT NULL;

//ALTER TABLE orfapen
//ALTER COLUMN Cod_Inte INT;
