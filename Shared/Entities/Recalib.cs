using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa los registros de calibración de instrumentos de medición.
    /// Tabla: RECALIB
    /// Migrada desde: Legacy_VB6/ArchivosDat/recalib.dat
    /// </summary>
    [Table("RECALIB")]
    public class Recalib : ITimestampEntity, IAuditableEntity
    {
        // ===================================================================
        // CLAVE PRIMARIA
        // ===================================================================

        /// <summary>
        /// ID autogenerado (agregado para EF Core)
        /// </summary>
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        // ===================================================================
        // CAMPOS DE NEGOCIO
        // ===================================================================

        /// <summary>
        /// Número de registro de calibración
        /// Posición en .dat: bytes 1-4 (CVS - binario a long)
        /// </summary>
        [Column("RegCalibr_Nro")]
        public int? RegCalibrNro { get; set; }

        /// <summary>
        /// Fecha de calibración
        /// Posición en .dat: bytes 5-6 (CVI - binario a integer)
        /// </summary>
        [Column("FeCalibracion")]
        public DateTime? FeCalibracion { get; set; }

        /// <summary>
        /// Fecha de próxima calibración
        /// Posición en .dat: bytes 7-8 (CVI - binario a integer)
        /// </summary>
        [Column("FeProxCalibr")]
        public DateTime? FeProxCalibr { get; set; }

        /// <summary>
        /// Número de instrumento (IC_NRO)
        /// Posición en .dat: bytes 17-32 (16 caracteres)
        /// </summary>
        [Column("IC_NRO")]
        [StringLength(16)]
        public string? IcNro { get; set; }

        /// <summary>
        /// Patrón utilizado para la calibración
        /// Posición en .dat: bytes 33-48 (16 caracteres)
        /// </summary>
        [Column("PatronUtilizado")]
        [StringLength(16)]
        public string? PatronUtilizado { get; set; }

        /// <summary>
        /// Indicador si requirió calibración
        /// Posición en .dat: bytes 49-64 (16 caracteres)
        /// </summary>
        [Column("RequirioCalib")]
        [StringLength(16)]
        public string? RequirioCalib { get; set; }

        /// <summary>
        /// Indicador si requirió ajuste
        /// Posición en .dat: bytes 65-80 (16 caracteres)
        /// </summary>
        [Column("RequirioAjuste")]
        [StringLength(16)]
        public string? RequirioAjuste { get; set; }

        /// <summary>
        /// Observaciones del registro de calibración
        /// Campo de texto largo
        /// </summary>
        [Column("OBSERVACIONES")]
        public string? Observaciones { get; set; }

        /// <summary>
        /// Responsable de la calibración
        /// </summary>
        [Column("RESPONSABLE")]
        [StringLength(32)]
        public string? Responsable { get; set; }

        /// <summary>
        /// Código interno del instrumento de medición (relación con INSTRUME.CodIntMedIdent)
        /// </summary>
        [Column("CodIntMedIdent")]
        public int? CodIntMedIdent { get; set; }

        /// <summary>
        /// Estado del registro
        /// 0 = Activo, -1 = Eliminado
        /// </summary>
        [Column("Status")]
        public short Status { get; set; } = 0;

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
 * ESTRUCTURA ORIGINAL DEL ARCHIVO LEGACY (recalib.dat)
 * =====================================================
 * 
 * Posiciones de campos en el archivo binario:
 * - Bytes 1-4:   RegCalibr_Nro (CVS - conversión binaria a long)
 * - Bytes 5-6:   FeCalibracion (CVI - conversión binaria a integer, representa fecha)
 * - Bytes 7-8:   FeProxCalibr (CVI - conversión binaria a integer, representa fecha)
 * - Bytes 9-16:  [No especificado - posiblemente padding]
 * - Bytes 17-32: IC_NRO (16 caracteres)
 * - Bytes 33-48: PatronUtilizado (16 caracteres)
 * - Bytes 49-64: RequirioCalib (16 caracteres)
 * - Bytes 65-80: RequirioAjuste (16 caracteres)
 * 
 * NOTAS DE MIGRACIÓN:
 * ===================
 * 1. CVS: Función VB6 que convierte 4 bytes binarios a un número long (Single)
 * 2. CVI: Función VB6 que convierte 2 bytes binarios a un integer
 * 3. Las fechas en VB6 están almacenadas como integers (días desde una fecha base)
 * 4. Los campos de texto pueden contener caracteres nulos que deben limpiarse
 * 5. CodIntMedIdent y OBSERVACIONES no están en el archivo .dat original,
 *    son campos adicionales para la base de datos SQL
 * 
 * RELACIONES:
 * ===========
 * - CodIntMedIdent -> INSTRUME.CodIntMedIdent (relación con instrumentos)
 * 
 * MAPEO FRONT ↔ BACK ↔ FRONT:
 * ===========================
 * UI (Blazor) ← DTO ← Entity (EF Core) ← SQL Server
 * - FeCalibracion: DatePicker → DateTime → DATETIME
 * - FeProxCalibr: DatePicker → DateTime → DATETIME
 * - IcNro: Input text → string → VARCHAR(16)
 * - PatronUtilizado: Input text → string → VARCHAR(16)
 * - RequirioCalib: Checkbox/Select → string → VARCHAR(16)
 * - RequirioAjuste: Checkbox/Select → string → VARCHAR(16)
 * - Observaciones: Textarea → string → VARCHAR(MAX)
 * - Responsable: Input text → string → VARCHAR(32)
 * - CodIntMedIdent: Input number → int → INT
 * - Status: short → SMALLINT (0 = Activo, -1 = Eliminado)
 * 
 * =====================================================
 * SCRIPT SQL PARA CREAR LA TABLA
 * =====================================================
 * 
 * CREATE TABLE [dbo].[RECALIB](
 *     [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
 *     [RegCalibr_Nro] INT NULL,
 *     [FeCalibracion] DATETIME NULL,
 *     [FeProxCalibr] DATETIME NULL,
 *     [IC_NRO] VARCHAR(16) NULL,
 *     [PatronUtilizado] VARCHAR(16) NULL,
 *     [RequirioCalib] VARCHAR(16) NULL,
 *     [RequirioAjuste] VARCHAR(16) NULL,
 *     [OBSERVACIONES] VARCHAR(MAX) NULL,
 *     [RESPONSABLE] VARCHAR(32) NULL,
 *     [CodIntMedIdent] INT NULL,
 *     [Status] SMALLINT NOT NULL DEFAULT 0,
 *     [AddRecord] DATETIME NULL,
 *     [LastUpdate] DATETIME NULL,
 *     [CodiEmprNet] INT NULL,
 *     [NumUsuar] INT NULL
 * );
 * GO
 * 
 * -- Índice para búsquedas por instrumento
 * CREATE INDEX IX_RECALIB_CodIntMedIdent ON [dbo].[RECALIB] ([CodIntMedIdent]);
 * GO

 */