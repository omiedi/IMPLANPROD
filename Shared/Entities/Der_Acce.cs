using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa los derechos de acceso de usuarios a aplicaciones.
    /// Mapea a la tabla DER_ACCE en la base de datos.
    /// </summary>
    [Table("DER_ACCE")]
    public class DerAcce : IAuditableEntity
    {
        /// <summary>
        /// Identificador interno autoincremental (IDENTITY) y clave primaria.
        /// Se agrega para usar una PK numérica por rendimiento y mantener los campos originales como datos de negocio.
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Número de usuario (obligatorio, debe ser mayor a 0)
        /// </summary>
        [Column("NUM_USER")]
        [Required(ErrorMessage = "El número de usuario es obligatorio")]
        [Range(1, short.MaxValue, ErrorMessage = "El número de usuario debe ser mayor a 0")]
        public short NumUser { get; set; }

        /// <summary>
        /// Código de aplicación (obligatorio, no puede estar vacío)
        /// </summary>
        [Column("COD_APLI")]
        [Required(ErrorMessage = "El código de aplicación es obligatorio")]
        [MaxLength(32)]
        [MinLength(1, ErrorMessage = "El código de aplicación no puede estar vacío")]
        public string CodApli { get; set; } = string.Empty;

        /// <summary>
        /// Derecho de lectura (read) - Por defecto 0
        /// </summary>
        [Column("DER_READ")]
        public short DerRead { get; set; } = 0;

        /// <summary>
        /// Derecho de escritura (write) - Por defecto 0
        /// </summary>
        [Column("DER_WRITE")]
        public short DerWrite { get; set; } = 0;

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
 * INSTRUCCIONES SQL PARA MODIFICAR LA TABLA DER_ACCE:
 * 
 * -- Agregar campo ID IDENTITY como clave primaria
 * ALTER TABLE DER_ACCE ADD Id INT IDENTITY(1,1) NOT NULL;
 * ALTER TABLE DER_ACCE ADD CONSTRAINT PK_DER_ACCE_Id PRIMARY KEY (Id);
 * 
 * -- Modificar campos existentes para hacerlos NOT NULL con valores por defecto
 * UPDATE DER_ACCE SET NUM_USER = 0 WHERE NUM_USER IS NULL;
 * ALTER TABLE DER_ACCE ALTER COLUMN NUM_USER SMALLINT NOT NULL;
 * 
 * UPDATE DER_ACCE SET COD_APLI = '' WHERE COD_APLI IS NULL;
 * ALTER TABLE DER_ACCE ALTER COLUMN COD_APLI NVARCHAR(32) NOT NULL;
 * 
 * UPDATE DER_ACCE SET DER_READ = 0 WHERE DER_READ IS NULL;
 * ALTER TABLE DER_ACCE ALTER COLUMN DER_READ SMALLINT NOT NULL;
 * 
 * UPDATE DER_ACCE SET DER_WRITE = 0 WHERE DER_WRITE IS NULL;
 * ALTER TABLE DER_ACCE ALTER COLUMN DER_WRITE SMALLINT NOT NULL;
 * 
 * 
 * -- Agregar campos de auditoría
 ALTER TABLE DER_ACCE ADD CODIEMPRNET INT NULL;
 ALTER TABLE DER_ACCE ADD NUMUSUAR INT NULL;
 ALTER TABLE DER_ACCE ADD ADDRECORD DATETIME2 NULL;
 ALTER TABLE DER_ACCE ADD LASTUPDATE DATETIME2 NULL;

  ALTER TABLE DER_ACCE ADD Id INT IDENTITY(1,1) NOT NULL;
  ALTER TABLE DER_ACCE ADD CONSTRAINT PK_DER_ACCE_Id PRIMARY KEY (Id);

* VALORES POR DEFECTO
  UPDATE dbo.DER_ACCE
SET
    ADDRECORD   = ISNULL(ADDRECORD,  GETDATE()),
    LASTUPDATE = ISNULL(LASTUPDATE, GETDATE()),
    CODIEMPRNET = ISNULL(CODIEMPRNET, 101),
    NUMUSUAR     = ISNULL(NUMUSUAR, 10)
WHERE
    ADDRECORD IS NULL
 OR LASTUPDATE IS NULL
 OR CODIEMPRNET IS NULL
 OR NUMUSUAR IS NULL;
 * 
 * -- Nota: La tabla original tiene un campo duplicado "DER-WRITE" que se ignora
 * -- Solo se usa el campo DER_WRITE (sin guión)
 */
