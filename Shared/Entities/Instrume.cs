using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad INSTRUME - Instrumentos de medición y calibración
    /// Registra los instrumentos utilizados en el sistema de calidad,
    /// incluyendo datos de identificación, calibración y ubicación.
    /// </summary>
    [Table("INSTRUME")]
    public class Instrume : ITimestampEntity, IAuditableEntity
    {
        // ===================================================================
        // CLAVE PRIMARIA
        // ===================================================================

        /// <summary>
        /// Identificador único auto-incremental
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("Id")]
        public int Id { get; set; }

        // ===================================================================
        // CAMPOS DE DATOS
        // ===================================================================

        /// <summary>
        /// Código identificador del instrumento
        /// </summary>
        [Column("Codigo")]
        [MaxLength(24)]
        public string? Codigo { get; set; }

        /// <summary>
        /// Descripción del instrumento
        /// </summary>
        [Column("Descripcion")]
        [MaxLength(64)]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Número de grupo al que pertenece el instrumento
        /// </summary>
        [Column("NroGrupo")]
        public int? NroGrupo { get; set; }

        /// <summary>
        /// Tipo de instrumento
        /// </summary>
        [Column("Tipo")]
        [MaxLength(24)]
        public string? Tipo { get; set; }

        /// <summary>
        /// Marca del instrumento
        /// </summary>
        [Column("Marca")]
        [MaxLength(48)]
        public string? Marca { get; set; }

        /// <summary>
        /// Modelo del instrumento
        /// </summary>
        [Column("Modelo")]
        [MaxLength(48)]
        public string? Modelo { get; set; }

        /// <summary>
        /// Número de serie del instrumento
        /// </summary>
        [Column("NroSerie")]
        [MaxLength(36)]
        public string? NroSerie { get; set; }

        /// <summary>
        /// Rango de medición del instrumento
        /// </summary>
        [Column("Rango")]
        [MaxLength(48)]
        public string? Rango { get; set; }

        /// <summary>
        /// Posición asignada al instrumento
        /// </summary>
        [Column("Posicion")]
        [MaxLength(48)]
        public string? Posicion { get; set; }

        /// <summary>
        /// Ubicación física del instrumento
        /// </summary>
        [Column("Ubicacion")]
        [MaxLength(48)]
        public string? Ubicacion { get; set; }

        /// <summary>
        /// Intervalo de calibración
        /// </summary>
        [Column("IntervCalibra")]
        [MaxLength(48)]
        public string? IntervCalibra { get; set; }

        /// <summary>
        /// Código interno del instrumento de identificación relacionado
        /// </summary>
        [Column("CodIntMedIdent")]
        public int? CodIntMedIdent { get; set; }

        /// <summary>
        /// Código interno del instrumento de calibración relacionado
        /// </summary>
        [Column("CodIntMedCalib")]
        public int? CodIntMedCalib { get; set; }

        /// <summary>
        /// Observaciones / notas extendidas
        /// </summary>
        [Column("Observaciones")]
        public string? Observaciones { get; set; }

        /// <summary>
        /// Estado del registro: 0 = activo, -1 = eliminado
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
        /// Número de usuario activo que creó/modificó el registro
        /// Se corresponde con Usuarios.IdFlexoft (numeración del sistema legacy)
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
============================================================
SCRIPT SQL - Crear tabla INSTRUME en SQL Server
============================================================
Ejecutar UNA SOLA VEZ en la base de datos de destino.

CREATE TABLE [dbo].[INSTRUME] (
    [Id]             [int]            IDENTITY(1,1)  NOT NULL,
    [Codigo]         [nvarchar](24)   NULL,
    [Descripcion]    [nvarchar](64)   NULL,
    [NroGrupo]       [int]            NULL,
    [Tipo]           [nvarchar](24)   NULL,
    [Marca]          [nvarchar](48)   NULL,
    [Modelo]         [nvarchar](48)   NULL,
    [NroSerie]       [nvarchar](36)   NULL,
    [Rango]          [nvarchar](48)   NULL,
    [Posicion]       [nvarchar](48)   NULL,
    [Ubicacion]      [nvarchar](48)   NULL,
    [IntervCalibra]  [nvarchar](48)   NULL,
    [CodIntMedIdent] [int]            NULL,
    [CodIntMedCalib] [int]            NULL,
    [Observaciones]  [nvarchar](MAX)  NULL,
    [Status]         [smallint]       NOT NULL  DEFAULT 0,
    [AddRecord]      [datetime]       NULL,
    [LastUpdate]     [datetime]       NULL,
    [CodiEmprNet]    [int]            NULL,
    [NumUsuar]       [int]            NULL,
    CONSTRAINT [PK_INSTRUME] PRIMARY KEY CLUSTERED ([Id] ASC)
);
GO

CREATE NONCLUSTERED INDEX [IX_INSTRUME_Codigo]
    ON [dbo].[INSTRUME] ([Codigo] ASC)
    WHERE [Codigo] IS NOT NULL;
GO

CREATE NONCLUSTERED INDEX [IX_INSTRUME_Status]
    ON [dbo].[INSTRUME] ([Status] ASC);
GO
============================================================
*/
