using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2022.M03.Main;
using DocumentFormat.OpenXml.Wordprocessing;
using IMPLANPROD.Shared.Entities;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad SOLICITUDES_AVANCES - Historial de avances de cada solicitud.
    /// Cada cambio de estado genera un nuevo registro; no se sobrescriben.
    /// </summary>
    [Table("SOLICITUDES_AVANCES")]
    public class SolicitudAvance : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador unico autoincremental del avance
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("IdAvance")]
        public int Id { get; set; }

        /// <summary>
        /// Identificador de la solicitud a la que pertenece el avance
        /// </summary>
        [Column("IdSolicitud")]
        public int SolicitudId { get; set; }

        /// <summary>
        /// Fecha en que se registro el avance
        /// </summary>
        [Column("Fecha")]
        public DateTime Fecha { get; set; }

        /// <summary>
        /// Estado reflejado en este avance
        /// </summary>
        [Column("Estado")]
        [MaxLength(50)]
        public string Estado { get; set; } = string.Empty;

        /// <summary>
        /// Comentario asociado al avance
        /// </summary>
        [Column("Comentario")]
        public string Comentario { get; set; } = string.Empty;

        /// <summary>
        /// IdFlexoft del usuario que registro el avance
        /// </summary>
        [Column("Usuario")]
        public int Usuario { get; set; }

        /// <summary>
        /// Navegacion a la solicitud padre
        /// </summary>
        public Solicitud? Solicitud { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMATICOS - Implementacion de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creacion del registro
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de ultima actualizacion del registro
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementacion de IAuditableEntity.CodiEmprNet
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementacion de IAuditableEntity.NumUsuar
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

        /// <summary>
        /// Código de empresa legacy, default 0
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; } = 0;
    }
}

/*
-- SQL ORIGINAL / CREACION:
-- Ver SQL/Create_Solicitudes.sql para el script completo de creacion
-- de las tablas SOLICITUDES y SOLICITUDES_AVANCES e indices.
*/
/*
- ===================================================================
--CREACION DE TABLAS DE SEGUIMIENTO DE SOLICITUDES
-- IMPORTANTE: Ejecutar en SQL Server Management Studio antes de usar
-- el modulo. NO se utilizan migrations automaticas.
-- ===================================================================

IF OBJECT_ID('dbo.SOLICITUDES', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.SOLICITUDES (
        IdSolicitud INT IDENTITY(1,1) NOT NULL,
        FechaIngreso DATETIME2 NOT NULL CONSTRAINT DF_SOLICITUDES_FechaIngreso DEFAULT(SYSDATETIME()),
        UsuarioSolicitante INT NOT NULL,
        Modulo NVARCHAR(80) NOT NULL,
        TipoSolicitud NVARCHAR(50) NOT NULL,
        Nivel NVARCHAR(50) NOT NULL,
        Titulo NVARCHAR(255) NOT NULL,
        Detalle NVARCHAR(MAX) NOT NULL,
        Estado NVARCHAR(50) NOT NULL CONSTRAINT DF_SOLICITUDES_Estado DEFAULT('Ingresado'),
        FechaUltimaActualizacion DATETIME2 NOT NULL CONSTRAINT DF_SOLICITUDES_FechaUltimaActualizacion DEFAULT(SYSDATETIME()),
        UsuarioUltimaActualizacion INT NULL,
        FechaCierre DATETIME2 NULL,
        UsuarioCierre INT NULL,
        AddRecord DATETIME2 NOT NULL CONSTRAINT DF_SOLICITUDES_AddRecord DEFAULT(SYSDATETIME()),
        LastUpdate DATETIME2 NOT NULL CONSTRAINT DF_SOLICITUDES_LastUpdate DEFAULT(SYSDATETIME()),
        CodiEmpr INT NOT NULL CONSTRAINT DF_SOLICITUDES_CodiEmpr DEFAULT(0),
        CodiEmprNet INT NOT NULL CONSTRAINT DF_SOLICITUDES_CodiEmprNet DEFAULT(0),
        NumUsuar INT NOT NULL CONSTRAINT DF_SOLICITUDES_NumUsuar DEFAULT(0),
        CONSTRAINT PK_SOLICITUDES PRIMARY KEY (IdSolicitud)
    );
END;

IF OBJECT_ID('dbo.SOLICITUDES_AVANCES', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.SOLICITUDES_AVANCES (
        IdAvance INT IDENTITY(1,1) NOT NULL,
        IdSolicitud INT NOT NULL,
        Fecha DATETIME2 NOT NULL CONSTRAINT DF_SOLICITUDES_AVANCES_Fecha DEFAULT(SYSDATETIME()),
        Estado NVARCHAR(50) NOT NULL,
        Comentario NVARCHAR(MAX) NOT NULL,
        Usuario INT NOT NULL,
        AddRecord DATETIME2 NOT NULL CONSTRAINT DF_SOLICITUDES_AVANCES_AddRecord DEFAULT(SYSDATETIME()),
        LastUpdate DATETIME2 NOT NULL CONSTRAINT DF_SOLICITUDES_AVANCES_LastUpdate DEFAULT(SYSDATETIME()),
        CodiEmpr INT NOT NULL CONSTRAINT DF_SOLICITUDES_AVANCES_CodiEmpr DEFAULT(0),
        CodiEmprNet INT NOT NULL CONSTRAINT DF_SOLICITUDES_AVANCES_CodiEmprNet DEFAULT(0),
        NumUsuar INT NOT NULL CONSTRAINT DF_SOLICITUDES_AVANCES_NumUsuar DEFAULT(0),
        CONSTRAINT PK_SOLICITUDES_AVANCES PRIMARY KEY (IdAvance),
        CONSTRAINT FK_SOLICITUDES_AVANCES_SOLICITUD FOREIGN KEY (IdSolicitud) REFERENCES dbo.SOLICITUDES(IdSolicitud) ON DELETE CASCADE
    );
END;

-- ===================================================================
--INDICES RECOMENDADOS
-- ===================================================================

IF NOT EXISTS(SELECT 1 FROM sys.indexes WHERE name = 'IX_SOLICITUDES_CodiEmprNet')
    CREATE NONCLUSTERED INDEX IX_SOLICITUDES_CodiEmprNet ON dbo.SOLICITUDES(CodiEmprNet);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SOLICITUDES_UsuarioSolicitante')
    CREATE NONCLUSTERED INDEX IX_SOLICITUDES_UsuarioSolicitante ON dbo.SOLICITUDES(UsuarioSolicitante);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SOLICITUDES_FechaIngreso')
    CREATE NONCLUSTERED INDEX IX_SOLICITUDES_FechaIngreso ON dbo.SOLICITUDES(FechaIngreso);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SOLICITUDES_Estado')
    CREATE NONCLUSTERED INDEX IX_SOLICITUDES_Estado ON dbo.SOLICITUDES(Estado);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SOLICITUDES_Modulo')
    CREATE NONCLUSTERED INDEX IX_SOLICITUDES_Modulo ON dbo.SOLICITUDES(Modulo);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SOLICITUDES_ComboFiltros')
    CREATE NONCLUSTERED INDEX IX_SOLICITUDES_ComboFiltros ON dbo.SOLICITUDES(FechaIngreso, Modulo, TipoSolicitud, Nivel, Estado);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SOLICITUDES_AVANCES_IdSolicitud')
    CREATE NONCLUSTERED INDEX IX_SOLICITUDES_AVANCES_IdSolicitud ON dbo.SOLICITUDES_AVANCES(IdSolicitud);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SOLICITUDES_AVANCES_Fecha')
    CREATE NONCLUSTERED INDEX IX_SOLICITUDES_AVANCES_Fecha ON dbo.SOLICITUDES_AVANCES(Fecha);

*/