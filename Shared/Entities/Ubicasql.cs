using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("UBICASQL")]
    public class Ubicasql : ITimestampEntity, IAuditableEntity
    {
        [Key]
        //[DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        // NOTA: El campo Id en la base de datos debe ser INT (Int32) para evitar errores de cast
        // Ejecutar: ALTER TABLE dbo.UBICASQL ALTER COLUMN Id INT NOT NULL;
        public int Id { get; set; }

        [Column("Cod_Inte")]
        public int? CodInte { get; set; }

        [Column("CodiEmpr")]
        public int CodiEmpr { get; set; } = 0;

        [Column("CODIUBICA")]
        [MaxLength(32)]
        public string? Codiubica { get; set; }

        [Column("CANTIDAD")]
        public decimal? Cantidad { get; set; }

        [Column("OBSERVACION")]
        [MaxLength(64)]
        public string? Observacion { get; set; }

        [Column("Nume_Pro")]
        public int? NumePro { get; set; }

        [Column("MARBORRA")]
        public short? Marborra { get; set; }

        [Column("Status")]
        public short? Status { get; set; }

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
SQL Server Management Studio (T-SQL) - Modificaciones sugeridas para UBICASQL

1) Agregar PK identity

ALTER TABLE dbo.UBICASQL
ADD Id INT IDENTITY(1,1) NOT NULL;

ALTER TABLE dbo.UBICASQL
ADD CONSTRAINT PK_UBICASQL_Id PRIMARY KEY (Id);

2) Cambiar CodiEmpr de SMALLINT a INT y forzar default 0

ALTER TABLE dbo.UBICASQL
ALTER COLUMN CodiEmpr INT NULL;

UPDATE dbo.UBICASQL
SET CodiEmpr = 0
WHERE CodiEmpr IS NULL;

ALTER TABLE dbo.UBICASQL
ADD CONSTRAINT DF_UBICASQL_CodiEmpr DEFAULT (0) FOR CodiEmpr;

3) Cambiar CANTIDAD de FLOAT a DECIMAL (ajustar precisión/escala según negocio)

ALTER TABLE dbo.UBICASQL
ALTER COLUMN CANTIDAD DECIMAL(18, 4) NULL;

4) Agregar campos de auditoría

ALTER TABLE dbo.UBICASQL ADD AddRecord DATETIME NULL;
ALTER TABLE dbo.UBICASQL ADD LastUpdate DATETIME NULL;
ALTER TABLE dbo.UBICASQL ADD CodiEmprNet INT NULL;
ALTER TABLE dbo.UBICASQL ADD NumUsuar INT NULL;

5) Agregar campo Status (SHORT)

ALTER TABLE dbo.UBICASQL ADD Status SMALLINT NULL;

UPDATE dbo.UBICASQL
SET Status = 0
WHERE Status IS NULL;


ALTER TABLE dbo.UBICASQL
ADD Status short  NULL;

ALTER TABLE dbo.UBICASQL
ADD CONSTRAINT DF_UBICASQL_Status DEFAULT (0) FOR Status;

UPDATE  UBICASQL SET STATUS = 0
ALTER TABLE dbo.UBICASQL



*/
