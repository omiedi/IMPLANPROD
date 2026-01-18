using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa los indicadores de aplicaciones del sistema.
    /// Mapea a la tabla INDI_APLI en la base de datos.
    /// </summary>
    [Table("INDI_APLI")]
    public class IndiApli : IAuditableEntity
    {
        /// <summary>
        /// Identificador interno autoincremental (IDENTITY) y clave primaria.
        /// Se agrega para usar una PK numérica por rendimiento y mantener los campos originales como datos de negocio.
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Código de aplicación (obligatorio, no puede estar vacío)
        /// </summary>
        [Column("COD_APLI")]
        [Required(ErrorMessage = "El código de aplicación es obligatorio")]
        [MaxLength(32)]
        [MinLength(1, ErrorMessage = "El código de aplicación no puede estar vacío")]
        public string CodApli { get; set; } = string.Empty;

        /// <summary>
        /// Descripción de la aplicación
        /// </summary>
        [Column("DES_APLI")]
        [MaxLength(128)]
        public string? DesApli { get; set; }

        // ===== CAMPOS DE AUDITORÍA =====

        /// <summary>
        /// Código de empresa .NET para auditoría
        /// </summary>
        [Column("CODIEMPRNET")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario que realizó la operación
        /// </summary>
        [Column("NUMUSUAR")]
        public int? NumUsuar { get; set; }

        /// <summary>
        /// Fecha de creación del registro
        /// </summary>
        [Column("ADDRECORD")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// </summary>
        [Column("LASTUPDATE")]
        public DateTime? LastUpdate { get; set; }
    }
}

/*
 * INSTRUCCIONES SQL PARA CREAR LA TABLA INDI_APLI:
 * 
 * -- Crear la tabla INDI_APLI
 * CREATE TABLE [dbo].[INDI_APLI](
 *     [Id] [int] IDENTITY(1,1) NOT NULL,
 *     [COD_APLI] [nvarchar](32) NOT NULL,
 *     [DES_APLI] [nvarchar](128) NULL,
 *     [CODIEMPRNET] [INT] NULL,
 *     [NUMUSUAR] [int] NULL,
 *     [ADDRECORD] [datetime2] NULL,
 *     [LASTUPDATE] [datetime2] NULL,
 *     CONSTRAINT [PK_INDI_APLI_Id] PRIMARY KEY CLUSTERED ([Id] ASC)
 * ) ON [PRIMARY];
 * GO
 * 
 * -- Agregar constraint para validar COD_APLI no vacío
 * ALTER TABLE INDI_APLI ADD CONSTRAINT CK_INDI_APLI_CodApli CHECK (LEN(COD_APLI) > 0);
 * GO
 * 
 * -- Crear índice único en COD_APLI
 * CREATE UNIQUE NONCLUSTERED INDEX IX_INDI_APLI_CodApli 
 * ON [dbo].[INDI_APLI] ([COD_APLI] ASC);
 * GO
 * 
 * -- Crear relación con DER_ACCE
 * ALTER TABLE [dbo].[DER_ACCE]
 * ADD CONSTRAINT FK_DER_ACCE_INDI_APLI 
 * FOREIGN KEY ([COD_APLI]) REFERENCES [dbo].[INDI_APLI]([COD_APLI])
 * ON DELETE NO ACTION
 * ON UPDATE NO ACTION;
 * GO
 * 
 * -- Actualizar registros existentes con valores por defecto para campos de auditoría
 * UPDATE dbo.INDI_APLI
 * SET
 *     ADDRECORD   = ISNULL(ADDRECORD,  GETDATE()),
 *     LASTUPDATE  = ISNULL(LASTUPDATE, GETDATE()),
 *     CODIEMPRNET = ISNULL(CODIEMPRNET, 101),
 *     NUMUSUAR    = ISNULL(NUMUSUAR, 10)
 * WHERE
 *     ADDRECORD IS NULL
 *  OR LASTUPDATE IS NULL
 *  OR CODIEMPRNET IS NULL
 *  OR NUMUSUAR IS NULL;
 * GO
 * 
 * ***** TUVE UNOS PROBLEMAS QUE LOS SOLUCIONE ASI   ****************
 * 
 * SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'INDI_APLI';

Eliminar el índice dependiente

DROP INDEX IX_INDI_APLI_CodApli ON dbo.INDI_APLI;
GO

Asegurar datos válidos (NULL o vacío)
UPDATE dbo.INDI_APLI
SET COD_APLI = '  '
WHERE COD_APLI IS NULL OR LTRIM(RTRIM(COD_APLI)) = '';
GO

Modificar la columna (ahora sí)
ALTER TABLE dbo.INDI_APLI
ALTER COLUMN COD_APLI NVARCHAR(32) NOT NULL;
GO

Recrear el índice
CREATE INDEX IX_INDI_APLI_CodApli
ON dbo.INDI_APLI (COD_APLI);
GO


 */
