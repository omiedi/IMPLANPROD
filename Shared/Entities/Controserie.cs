using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("CONTROSERIE")]
    public class Controserie : ITimestampEntity, IAuditableEntity
    {
        [Key]
        [Column("Id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("CONTROSERIE_ID")]
        public int ControserieId { get; set; }

        [Column("Cod_Inte")]
        public int? CodInte { get; set; }

        [Column("CodiEmpr")]
        public int CodiEmpr { get; set; } = 0;

        [Column("Cod_Exte")]
        [StringLength(24)]
        public string? CodExte { get; set; }

        [Column("Descrip")]
        [StringLength(64)]
        public string? Descrip { get; set; }

        [Column("IdenLote")]
        [StringLength(12)]
        public string? IdenLote { get; set; }

        [Column("NumeroSerie")]
        [StringLength(24)]
        public string? NumeroSerie { get; set; }

        [Column("Observaciones")]
        [StringLength(48)]
        public string? Observaciones { get; set; }

        [Column("NumeroSerieOrigen")]
        [StringLength(24)]
        public string? NumeroSerieOrigen { get; set; }

        [Column("CodInte_DestinoConsumo")]
        public int? CodInteDestinoConsumo { get; set; }

        [Column("CodExte_DestinoConsumo")]
        [StringLength(24)]
        public string? CodExteDestinoConsumo { get; set; }

        [Column("Status")]
        public short? Status { get; set; } = 0;

        [Column("ORDEN")]
        public int? Orden { get; set; }

        [Column("MARBORRA")]
        public int? Marborra { get; set; }

        [Column("IdCodAsignado")]
        public int? IdCodAsignado { get; set; }

        [Column("DepoMovi")]
        public int? DepoMovi { get; set; }

        [Column("CantIngre", TypeName = "decimal(18,4)")]
        public decimal? CantIngre { get; set; }

        [Column("CantSalid", TypeName = "decimal(18,4)")]
        public decimal? CantSalid { get; set; }

        [Column("FEREMOVI")]
        public DateTime? FeReMovi { get; set; }

        [Column("IDSUBOR")]
        [StringLength(12)]
        public string? IdSubOr { get; set; }

        [Column("UsoBruto", TypeName = "decimal(18,4)")]
        public decimal? UsoBruto { get; set; }

        [Column("EspIdentado")]
        public int? EspIdentado { get; set; }

        [Column("IDSUBOR2")]
        [StringLength(16)]
        public string? IdSubOr2 { get; set; }

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
SQL Server Management Studio (T-SQL) - Crear tabla CONTROSERIE

drop TABLE [dbo].[CONTROSERIE]
USE [arcon2]
GO
CREAR LA TABLA Y TRANSMITIR LOS DATOS
--SELECT * 
--INTO arcon2.dbo.CONTROSERIE --tabla destion
--FROM arcon.dbo.CONTROSERIE -- tabla origen

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

--CAMBIARLE EL NOMBRE AL CAMPO CONTROSERIE_ID A ID
EXEC sp_rename 
    'CONTROSERIE.CONTROSERIE_ID',
    'Id',
    'COLUMN';
--VUELVO A CREAR EL CAMPO
ALTER TABLE CONTROSERIE
ADD CONTROSERIE_ID INT NULL;
----
--LE PASO LOS MISMOS VALORES DE ID
UPDATE CONTROSERIE
SET CONTROSERIE_ID = Id;
--CREO LOS CAMPOS DE AUDITORIA
ALTER TABLE CONTROSERIE
ADD AddRecord DATETIME NULL,
    LastUpdate DATETIME NULL,
    CodiEmprNet INT NULL,
    NumUsuar INT NULL;
--PASO A INT
ALTER TABLE CONTROSERIE
ALTER COLUMN Cod_Inte INT NULL;
GO
--VALRO POR DEFECTO
UPDATE CONTROSERIE
SET CodiEmpr = 0
WHERE CodiEmpr IS NULL;
GO

ALTER TABLE CONTROSERIE
ALTER COLUMN CodiEmpr INT NOT NULL;
GO
--PASO A INT
ALTER TABLE CONTROSERIE
ALTER COLUMN CodInte_DestinoConsumo INT NULL;
GO
--PASO A INT
ALTER TABLE CONTROSERIE
ALTER COLUMN ORDEN INT NULL;
GO
--PASO A INT
ALTER TABLE CONTROSERIE
ALTER COLUMN MARBORRA INT NULL;
GO
--PASO A INT
ALTER TABLE CONTROSERIE
ALTER COLUMN DepoMovi INT NULL;
GO
--PASO A INT
ALTER TABLE CONTROSERIE
ALTER COLUMN EspIdentado INT NULL;
GO
--PASO A DECIMAL
ALTER TABLE CONTROSERIE
ALTER COLUMN CantIngre DECIMAL(18,4) NULL;
GO
--PASO A DECIMAL
ALTER TABLE CONTROSERIE
ALTER COLUMN CantSalid DECIMAL(18,4) NULL;
GO
--PASO A DECIMAL
ALTER TABLE CONTROSERIE
ALTER COLUMN UsoBruto DECIMAL(18,4) NULL;
GO



ALTER TABLE CONTROSERIE
ADD CONSTRAINT DF_CONTROSERIE_CodiEmpr DEFAULT(0) FOR CodiEmpr;
GO

ALTER TABLE CONTROSERIE
ADD CONSTRAINT DF_CONTROSERIE_Status DEFAULT(0) FOR Status;
GO

sp_help CONTROSERIE;


UPDATE CONTROSERIE
SET NUMUSUAR =6;

UPDATE CONTROSERIE
SET CODIEMPRNET =10;

UPDATE CONTROSERIE
SET AddRecord = GETDATE()

UPDATE CONTROSERIE
SET lastupdate = GETDATE()
*/
