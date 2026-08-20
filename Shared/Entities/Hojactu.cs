using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad HOJACTU - Hoja de Actualización / Control de Cambios
    /// </summary>
    [Table("HOJACTU")]
    public class Hojactu : ITimestampEntity, IAuditableEntity
    {
        // ===================================================================
        // CLAVE PRIMARIA
        // ===================================================================

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("Id")]
        public int Id { get; set; }

        // ===================================================================
        // CAMPOS DE NEGOCIO
        // ===================================================================

        /// <summary>
        /// Código de empresa — siempre 0 según reglas globales del sistema
        /// </summary>
        [Column("CodiEmpr")]
        public int CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Fecha de la actualización
        /// </summary>
        [Column("Fecha")]
        public DateTime? Fecha { get; set; }

        /// <summary>
        /// Número de hoja (obligatorio)
        /// </summary>
        [Required(ErrorMessage = "El campo Hoja es obligatorio.")]
        [Column("Hoja")]
        public int Hoja { get; set; }

        /// <summary>
        /// Descripción corta de la modificación realizada (obligatorio)
        /// </summary>
        [Required(ErrorMessage = "El campo Modificacion es obligatorio.")]
        [MaxLength(255, ErrorMessage = "Modificacion no puede superar los 255 caracteres.")]
        [Column("Modificacion")]
        public string Modificacion { get; set; } = null!;

        /// <summary>
        /// Descripción extendida de la actualización
        /// </summary>
        [MaxLength(1028, ErrorMessage = "Descripcion no puede superar los 1028 caracteres.")]
        [Column("Descripcion")]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Código interno del artículo/instrucción al que aplica la actualización
        /// Relacionado con INSTRUCI.Cod_Inte y MASTER.CodInte (obligatorio)
        /// </summary>
        [Required(ErrorMessage = "El campo Cod_Inte es obligatorio.")]
        [Column("Cod_Inte")]
        public int Cod_Inte { get; set; }

        /// <summary>
        /// Estado del registro — 0: activo (valor por defecto)
        /// </summary>
        [Column("Status")]
        public short Status { get; set; } = 0;

        /// <summary>
        /// Orden de visualización dentro del grupo de hojas de la instrucción
        /// </summary>
        [Column("Orden")]
        public int? Orden { get; set; }

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
        /// Número de usuario activo — se corresponde con Usuarios.IdFlexoft
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
============================================================
SCRIPT SQL - Crear tabla HOJACTU en SQL Server
============================================================
Ejecutar UNA SOLA VEZ en la base de datos de destino.

CREATE TABLE [dbo].[HOJACTU] (
    [Id]           [int]            IDENTITY(1,1)  NOT NULL,   -- PK auto-incremental
    [CodiEmpr]     [int]            NOT NULL        DEFAULT 0,  -- Código empresa (siempre 0)
    [Fecha]        [datetime]       NULL,                       -- Fecha de la actualización
    [Hoja]         [int]            NOT NULL,                   -- Número de hoja
    [Modificacion] [nvarchar](255)  NOT NULL,                   -- Descripción corta de modificación
    [Descripcion]  [nvarchar](1028) NULL,                       -- Descripción extendida
    [Cod_Inte]     [int]            NOT NULL,                   -- Código interno (ref. INSTRUCI.Cod_Inte)
    [Status]       [smallint]       NOT NULL        DEFAULT 0,  -- Estado: 0=activo
    [Orden]        [int]            NULL,                       -- Orden de visualización
    [AddRecord]    [datetime]       NULL,                       -- Fecha creación registro
    [LastUpdate]   [datetime]       NULL,                       -- Fecha última modificación
    [CodiEmprNet]  [int]            NULL,                       -- Código empresa (auditoría)
    [NumUsuar]     [int]            NULL,                       -- Usuario activo (ref. Usuarios.IdFlexoft)
    CONSTRAINT [PK_HOJACTU] PRIMARY KEY CLUSTERED ([Id] ASC)
);
GO

-- Índice por Cod_Inte para joins con INSTRUCI y MASTER
CREATE NONCLUSTERED INDEX [IX_HOJACTU_Cod_Inte]
    ON [dbo].[HOJACTU] ([Cod_Inte] ASC);
GO

-- Índice compuesto por Cod_Inte + Hoja para búsquedas de hojas de un documento
CREATE NONCLUSTERED INDEX [IX_HOJACTU_Cod_Inte_Hoja]
    ON [dbo].[HOJACTU] ([Cod_Inte] ASC, [Hoja] ASC);
GO

-- Índice por CodiEmpr para compatibilidad con filtros multi-empresa
CREATE NONCLUSTERED INDEX [IX_HOJACTU_CodiEmpr]
    ON [dbo].[HOJACTU] ([CodiEmpr] ASC);
GO
============================================================
*/
