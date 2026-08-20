using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que mapea la tabla SQL_QUERIES de SQLCTRL.
    /// Almacena consultas SQL parametrizables del sistema.
    /// </summary>
    [Table("SQL_QUERIES")]
    public class SQL_QUERIES : ITimestampEntity, IAuditableEntity
    {
        // ── Clave técnica autoincremental ─────────────────────────────────────
        [Key]
        [Column("Id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        // ── Campos originales de la tabla ─────────────────────────────────────
        /// <summary>
        /// Código de la consulta (texto 16 caracteres)
        /// </summary>
        [Column("Cod_Query")]
        [StringLength(16)]
        public string? CodQuery { get; set; }

        /// <summary>
        /// Descripción de la consulta (texto 254 caracteres)
        /// </summary>
        [Column("Des_Query")]
        [StringLength(254)]
        public string? DesQuery { get; set; }

        /// <summary>
        /// Tipo de consulta (entero)
        /// </summary>
        [Column("Tip_Query")]
        public int? TipQuery { get; set; }

        /// <summary>
        /// Parámetros de la consulta (varchar(max))
        /// </summary>
        [Column("Par_Query", TypeName = "varchar(max)")]
        public string? ParQuery { get; set; }

        /// <summary>
        /// Consulta SQL propiamente dicha (varchar(max))
        /// </summary>
        [Column("SQL_Query", TypeName = "varchar(max)")]
        public string? SqlQuery { get; set; }

        // ── Campo de baja lógica ──────────────────────────────────────────────
        /// <summary>
        /// Estado del registro. 0 = activo, -1 = dado de baja
        /// </summary>
        [Column("Status")]
        public short? Status { get; set; } = 0;

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

        // ===================================================================
        // CAMPOS DE AUDITORÍA - Implementación de IAuditableEntity
        // ===================================================================

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
SQL Server Management Studio (T-SQL) - SQL Original

USE [SQLCTRL]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[SQL_QUERIES](
    [Cod_Query] [varchar](16) NULL,
    [Des_Query] [varchar](254) NULL,
    [Tip_Query] [int] NULL,
    [Par_Query] [varchar](max) NULL,
    [SQL_Query] [varchar](max) NULL
)
GO

-- CAMBIOS APLICADOS PARA IMPLANPROD:
-- 1. Se agrega clave técnica autoincremental [Id]
-- 2. Se agrega campo [Status] para baja lógica
-- 3. Se agregan campos de auditoría: AddRecord, LastUpdate, CodiEmprNet, NumUsuar

ALTER TABLE [dbo].[SQL_QUERIES]
ADD [Id] [int] IDENTITY(1,1) NOT NULL;
GO

ALTER TABLE [dbo].[SQL_QUERIES]
ADD [Status] [smallint] NULL DEFAULT(0);
GO

ALTER TABLE [dbo].[SQL_QUERIES]
ADD [AddRecord] [datetime] NULL,
    [LastUpdate] [datetime] NULL,
    [CodiEmprNet] [int] NULL,
    [NumUsuar] [int] NULL;
GO

sp_help [SQL_QUERIES];
*/
