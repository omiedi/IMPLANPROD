using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("REPLACO")]
    public class Replaco : ITimestampEntity, IAuditableEntity
    {
        [Key]
        [Column("ID")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("NUME_CLI")]
        public int? NumeCli { get; set; }

        [Column("NUME_PROV")]
        public int? NumeProv { get; set; }

        [Required]
        [Column("COD_INTE")]
        public int CodInte { get; set; }

        [Column("REFERENCIA")]
        [StringLength(64)]
        public string? Referencia { get; set; }

        [Column("OBSERVACIONES")]
        [StringLength(256)]
        public string? Observaciones { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
CREATE TABLE [dbo].[REPLACO](
    [ID] INT IDENTITY(1,1) NOT NULL,
    [NUME_CLI] INT NULL,
    [NUME_PROV] INT NULL,
    [COD_INTE] INT NOT NULL,
    [REFERENCIA] VARCHAR(64) NULL,
    [OBSERVACIONES] VARCHAR(256) NULL,
    [AddRecord] DATETIME NULL,
    [LastUpdate] DATETIME NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,
    CONSTRAINT [PK_REPLACO] PRIMARY KEY CLUSTERED ([ID] ASC)
);
*/
