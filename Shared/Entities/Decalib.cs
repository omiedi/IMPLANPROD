using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("DECALIB")]
    public class Decalib : ITimestampEntity, IAuditableEntity
    {
        [Key]
        [Column("Id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("PATRON")]
        [StringLength(32)]
        public string? Patron { get; set; }

        [Column("MEDIDA", TypeName = "decimal(18,3)")]
        public decimal? Medida { get; set; }

        [Column("X1", TypeName = "decimal(18,3)")]
        public decimal? X1 { get; set; }

        [Column("X2", TypeName = "decimal(18,3)")]
        public decimal? X2 { get; set; }

        [Column("X3", TypeName = "decimal(18,3)")]
        public decimal? X3 { get; set; }

        [Column("X4", TypeName = "decimal(18,3)")]
        public decimal? X4 { get; set; }

        [Column("X5", TypeName = "decimal(18,3)")]
        public decimal? X5 { get; set; }

        [Column("E_prom", TypeName = "decimal(18,3)")]
        public decimal? EProm { get; set; }

        [Column("I_me", TypeName = "decimal(18,3)")]
        public decimal? IMe { get; set; }

        [Column("IdReCalib")]//es el id de la tabla recalib
        public int IdReCalib { get; set; }

        [Column("Status")]
        public short Status { get; set; } = 0;

        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; } = 0;

        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
CREATE TABLE [dbo].[DECALIB](
    [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [PATRON] VARCHAR(32) NULL,
    [MEDIDA] DECIMAL(18,3) NULL,
    [X1] DECIMAL(18,3) NULL,
    [X2] DECIMAL(18,3) NULL,
    [X3] DECIMAL(18,3) NULL,
    [X4] DECIMAL(18,3) NULL,
    [X5] DECIMAL(18,3) NULL,
    [E_prom] DECIMAL(18,3) NULL,
    [I_me] DECIMAL(18,3) NULL,
    [IdReCalib] INT NOT NULL,
    [Status] SMALLINT NOT NULL DEFAULT 0,
    [AddRecord] DATETIME NULL,
    [LastUpdate] DATETIME NULL,
    [CodiEmprNet] INT NULL DEFAULT 0,
    [NumUsuar] INT NULL
);
*/
