using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad de Transportes que mapea la tabla física TRANSPORTES.
    /// Incluye campos básicos de identificación y contacto del transportista.
    /// </summary>
    public class Transporte : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador interno Identity para EF y relaciones.
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Código del transportista (equivale a Codi_Tra). Se migra a int.
        /// </summary>
        [Display(Name = "Código")]
        public int? Codi_Tra { get; set; }

        /// <summary>
        /// Código de empresa (equivale a CodiEmpr). Se migra a int.
        /// </summary>
        [Display(Name = "Empresa")]
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Denominación o nombre corto (Deno_Tra)
        /// </summary>
        [MaxLength(128)]
        [Display(Name = "Nombre")]
        public string? Deno_Tra { get; set; }

        /// <summary>
        /// Razón Social (Raso_Tra)
        /// </summary>
        [MaxLength(128)]
        [Display(Name = "Razón Social")]
        public string? Raso_Tra { get; set; }

        /// <summary>
        /// Domicilio (Domi_Tra)
        /// </summary>
        [MaxLength(128)]
        [Display(Name = "Domicilio")]
        public string? Domi_Tra { get; set; }

        /// <summary>
        /// Localidad (Loca_Tra)
        /// </summary>
        [MaxLength(128)]
        [Display(Name = "Localidad")]
        public string? Loca_Tra { get; set; }

        /// <summary>
        /// CUIT (Cuit_Tra)
        /// </summary>
        [MaxLength(32)]
        [Display(Name = "CUIT")]
        public string? Cuit_Tra { get; set; }

        /// <summary>
        /// Posición de IVA (Piva_Tra) como short.
        /// </summary>
        [Display(Name = "Posición IVA")]
        public short? Piva_Tra { get; set; }

        /// <summary>
        /// Teléfono (Tele_Tra)
        /// </summary>
        [MaxLength(64)]
        [Display(Name = "Teléfono")]
        public string? Tele_Tra { get; set; }

        /// <summary>
        /// Fax (Faxi_Tra)
        /// </summary>
        [MaxLength(64)]
        public string? Faxi_Tra { get; set; }

        /// <summary>
        /// Email (Mail_Tra)
        /// </summary>
        [MaxLength(128)]
        [EmailAddress]
        public string? Mail_Tra { get; set; }

        /// <summary>
        /// Contacto (Ctac_Tra)
        /// </summary>
        [MaxLength(128)]
        [Display(Name = "Contacto")]
        public string? Ctac_Tra { get; set; }

        /// <summary>
        /// Provincia (Prov_Tra)
        /// </summary>
        [MaxLength(4)]
        public string? Prov_Tra { get; set; }

        /// <summary>
        /// Código Postal (CodPostal_Tra)
        /// </summary>
        [MaxLength(24)]
        public string? CodPostal_Tra { get; set; }

        /// <summary>
        /// Hora del transporte (HoraTransporte)
        /// </summary>
        [MaxLength(48)]
        public string? HoraTransporte { get; set; }

        /// <summary>
        /// Estado (0 activo, -1 inactivo)
        /// </summary>
        public short? Status { get; set; }

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

/*
-- SQL base original (comentado a pedido):
CREATE TABLE [dbo].[TRANSPORTES](
    [Codi_Tra] [smallint] NULL,
    [CodiEmpr] [smallint] NULL,
    [Deno_Tra] [nvarchar](128) NULL,
    [Raso_Tra] [nvarchar](128) NULL,
    [Domi_Tra] [nvarchar](128) NULL,
    [Loca_Tra] [nvarchar](128) NULL,
    [Cuit_Tra] [nvarchar](32) NULL,
    [Piva_Tra] [smallint] NULL,
    [Tele_Tra] [nvarchar](64) NULL,
    [Faxi_Tra] [nvarchar](64) NULL,
    [Mail_Tra] [nvarchar](128) NULL,
    [Ctac_Tra] [nvarchar](128) NULL,
    [Prov_Tra] [nvarchar](4) NULL,
    [CodPostal_Tra] [nvarchar](24) NULL,
    [HoraTransporte] [nvarchar](48) NULL
) ON [PRIMARY]
GO

-- Ajustes solicitados:
-- 1) Agregar columna Id Identity
IF COL_LENGTH('dbo.TRANSPORTES','Id') IS NULL
BEGIN
    ALTER TABLE dbo.TRANSPORTES ADD Id INT IDENTITY(1,1) NOT NULL;
    ALTER TABLE dbo.TRANSPORTES ADD CONSTRAINT PK_TRANSPORTES_Id PRIMARY KEY (Id);
END

-- 2) Convertir Codi_Tra de smallint a int
--    (crear columna temporal, copiar datos, eliminar y renombrar para entornos con datos)
IF EXISTS(SELECT 1 FROM sys.columns WHERE Name = 'Codi_Tra' AND Object_ID = Object_ID('dbo.TRANSPORTES'))
BEGIN
    ALTER TABLE dbo.TRANSPORTES ADD Codi_Tra_int INT NULL;
    UPDATE dbo.TRANSPORTES SET Codi_Tra_int = CAST(Codi_Tra AS INT);
    ALTER TABLE dbo.TRANSPORTES DROP COLUMN Codi_Tra;
    EXEC sp_rename 'dbo.TRANSPORTES.Codi_Tra_int', 'Codi_Tra', 'COLUMN';
END

-- 3) Convertir CodiEmpr de smallint a int
IF EXISTS(SELECT 1 FROM sys.columns WHERE Name = 'CodiEmpr' AND Object_ID = Object_ID('dbo.TRANSPORTES'))
BEGIN
    ALTER TABLE dbo.TRANSPORTES ADD CodiEmpr_int INT NULL;
    UPDATE dbo.TRANSPORTES SET CodiEmpr_int = CAST(CodiEmpr AS INT);
    ALTER TABLE dbo.TRANSPORTES DROP COLUMN CodiEmpr;
    EXEC sp_rename 'dbo.TRANSPORTES.CodiEmpr_int', 'CodiEmpr', 'COLUMN';
END

-- 4) Agregar columna Status short (smallint)
IF COL_LENGTH('dbo.TRANSPORTES','Status') IS NULL
BEGIN
    ALTER TABLE dbo.TRANSPORTES ADD Status SMALLINT NULL;
END
*/
