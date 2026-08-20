using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad para auditoría de archivos de sincronización procesados
    /// </summary>
    [Table("SyncProcesados")]
    public class SyncProcesado : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador único
        /// </summary>
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Nombre del archivo procesado
        /// </summary>
        [Required]
        [MaxLength(200)]
        public string Archivo { get; set; } = string.Empty;

        /// <summary>
        /// Tabla afectada (MASTER, MOVISTO, etc.)
        /// </summary>
        [Required]
        [MaxLength(50)]
        public string Tabla { get; set; } = string.Empty;

        /// <summary>
        /// Operación realizada (UPSERT, DELETE, etc.)
        /// </summary>
        [MaxLength(20)]
        public string? Operacion { get; set; }

        /// <summary>
        /// Fecha de procesamiento
        /// </summary>
        public DateTime FechaProcesamiento { get; set; }

        /// <summary>
        /// Estado del procesamiento (EXITOSO, ERROR, PENDIENTE)
        /// </summary>
        [Required]
        [MaxLength(20)]
        public string Estado { get; set; } = "PENDIENTE";

        /// <summary>
        /// Mensaje de error si hubo algún problema
        /// </summary>
        public string? Error { get; set; }

        /// <summary>
        /// Código interno del registro afectado (para MASTER sería Codint)
        /// </summary>
        public int? CodintAfectado { get; set; }

        /// <summary>
        /// Código de empresa por defecto
        /// </summary>
        [Column("Codiempr")]
        public short Codiempr { get; set; } = 0;

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
SQL PARA CREAR LA TABLA SyncProcesados

Este script crea la tabla usada por la entidad SyncProcesado.
La tabla registra cada archivo de sincronización procesado, su estado
y los datos básicos de auditoría del sistema.

CREATE TABLE [dbo].[SyncProcesados]
(
    [Id] INT IDENTITY(1,1) NOT NULL,
    [Archivo] NVARCHAR(200) NOT NULL,
    [Tabla] NVARCHAR(50) NOT NULL,
    [Operacion] NVARCHAR(20) NULL,
    [FechaProcesamiento] DATETIME2 NOT NULL,
    [Estado] NVARCHAR(20) NOT NULL CONSTRAINT [DF_SyncProcesados_Estado] DEFAULT ('PENDIENTE'),
    [Error] NVARCHAR(MAX) NULL,
    [CodintAfectado] INT NULL,
    [Codiempr] SMALLINT NOT NULL CONSTRAINT [DF_SyncProcesados_Codiempr] DEFAULT (0),
    [AddRecord] DATETIME2 NULL,
    [LastUpdate] DATETIME2 NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,

    CONSTRAINT [PK_SyncProcesados] PRIMARY KEY CLUSTERED ([Id] ASC)
);

CREATE INDEX [IX_SyncProcesados_Archivo]
    ON [dbo].[SyncProcesados] ([Archivo]);

CREATE INDEX [IX_SyncProcesados_Tabla_Estado]
    ON [dbo].[SyncProcesados] ([Tabla], [Estado]);

CREATE INDEX [IX_SyncProcesados_FechaProcesamiento]
    ON [dbo].[SyncProcesados] ([FechaProcesamiento]);
*/
