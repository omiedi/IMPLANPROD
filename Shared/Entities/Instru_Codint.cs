using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad INSTRU_CODINT - Relación entre instrucciones de ingeniería y artículos (MASTER)
    /// Vincula un código interno de MASTER con un código interno de INSTRUCI
    /// </summary>
    [Table("INSTRU_CODINT")]
    public class Instru_Codint : ITimestampEntity, IAuditableEntity
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
        /// Código interno del artículo en MASTER (MASTER.CodInte)
        /// </summary>
        [Column("Master_Cod_Inte")]
        public int Master_Cod_Inte { get; set; }

        /// <summary>
        /// Código interno de la instrucción en INSTRUCI (INSTRUCI.Cod_Inte)
        /// </summary>
        [Column("Instruci_Cod_Inte")]
        public int Instruci_Cod_Inte { get; set; }

        /// <summary>
        /// Destinatario / sector al que aplica esta relación
        /// </summary>
        [Column("Para")]
        [MaxLength(64)]
        public string? Para { get; set; }

        /// <summary>
        /// Estado del registro — 0: activo (valor por defecto)
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
        /// Número de usuario activo — se corresponde con Usuarios.IdFlexoft
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
============================================================
SCRIPT SQL - Crear tabla INSTRU_CODINT en SQL Server
============================================================
Ejecutar UNA SOLA VEZ en la base de datos de destino.

CREATE TABLE [dbo].[    ] (
    [Id]                [int]       IDENTITY(1,1)  NOT NULL,   -- PK auto-incremental
    [Master_Cod_Inte]   [int]       NOT NULL,                  -- Código interno artículo (ref. MASTER.CodInte)
    [Instruci_Cod_Inte] [int]       NOT NULL,                  -- Código interno instrucción (ref. INSTRUCI.Cod_Inte)
    [Para]              [nvarchar](64) NULL,                    -- Destinatario / sector
    [Status]            [smallint]  NOT NULL        DEFAULT 0,  -- Estado: 0=activo
    [AddRecord]         [datetime]  NULL,                       -- Fecha creación registro
    [LastUpdate]        [datetime]  NULL,                       -- Fecha última modificación
    [CodiEmprNet]       [int]       NULL,                       -- Código empresa (auditoría)
    [NumUsuar]          [int]       NULL,                       -- Usuario activo (ref. Usuarios.IdFlexoft)
    CONSTRAINT [PK_INSTRU_CODINT] PRIMARY KEY CLUSTERED ([Id] ASC)
);
GO

CREATE NONCLUSTERED INDEX [IX_INSTRU_CODINT_Master]
    ON [dbo].[INSTRU_CODINT] ([Master_Cod_Inte] ASC);
GO

CREATE NONCLUSTERED INDEX [IX_INSTRU_CODINT_Instruci]
    ON [dbo].[INSTRU_CODINT] ([Instruci_Cod_Inte] ASC);
GO
============================================================
*/
