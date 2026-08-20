using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("EmailConfigEmpresa")]
    public class EmailConfigEmpresa : ITimestampEntity, IAuditableEntity
    {
        public int Id { get; set; }

        [Required]
        [MaxLength(150)]
        public string SmtpHost { get; set; } = string.Empty;

        public int SmtpPort { get; set; } = 587;

        public bool UsarSsl { get; set; } = false;

        public bool UsarStartTls { get; set; } = true;

        [Required]
        [MaxLength(200)]
        public string UsuarioSmtp { get; set; } = string.Empty;

        [Required]
        public string ContrasenaEncriptada { get; set; } = string.Empty;

        [Required]
        [MaxLength(200)]
        public string EmailDesde { get; set; } = string.Empty;

        [Required]
        [MaxLength(200)]
        public string NombreDesde { get; set; } = string.Empty;

        public DateTime LastModified { get; set; }

        [Column("Codiempr")]
        public int? Codiempr { get; set; } = 0;

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
-- SCRIPT SQL PARA PRODUCCIÓN (ejecutar una sola vez):

IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='EmailConfigEmpresa' AND xtype='U')
BEGIN
    CREATE TABLE dbo.EmailConfigEmpresa
    (
        Id                  INT           IDENTITY(1,1) NOT NULL PRIMARY KEY,
        SmtpHost            NVARCHAR(150) NOT NULL,
        SmtpPort            INT           NOT NULL CONSTRAINT DF_ECE_SmtpPort     DEFAULT(587),
        UsarSsl             BIT           NOT NULL CONSTRAINT DF_ECE_UsarSsl      DEFAULT(0),
        UsarStartTls        BIT           NOT NULL CONSTRAINT DF_ECE_UsarStartTls DEFAULT(1),
        UsuarioSmtp         NVARCHAR(200) NOT NULL,
        ContrasenaEncriptada NVARCHAR(MAX) NOT NULL,
        EmailDesde          NVARCHAR(200) NOT NULL,
        NombreDesde         NVARCHAR(200) NOT NULL,
        LastModified        DATETIME2     NOT NULL CONSTRAINT DF_ECE_LastModified  DEFAULT(SYSDATETIME()),
        Codiempr            INT           NOT NULL CONSTRAINT DF_ECE_Codiempr     DEFAULT(0),
        AddRecord           DATETIME2     NULL     CONSTRAINT DF_ECE_AddRecord     DEFAULT(SYSDATETIME()),
        LastUpdate          DATETIME2     NULL     CONSTRAINT DF_ECE_LastUpdate    DEFAULT(SYSDATETIME()),
        CodiEmprNet         INT           NULL     CONSTRAINT DF_ECE_CodiEmprNet  DEFAULT(0),
        NumUsuar            INT           NULL     CONSTRAINT DF_ECE_NumUsuar     DEFAULT(0)
    );
END;
*/
