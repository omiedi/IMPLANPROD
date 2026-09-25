using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Tabla ANOTAPED - Anotaciones de pedidos estructuradas.
    /// Reemplaza el campo ntext PEDENCA.AnotaPed por registros individuales
    /// que permiten consultar y mostrar anotaciones de forma estructurada
    /// (productos con cantidades, observaciones por item y generales).
    /// Mantiene compatibilidad con el formato VB6 original de AGREGA_ANOTAPED.
    /// </summary>
    [Table("ANOTAPED")]
    public class Anotaped : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Identificador del grupo de registros registrados en cada update de anotación.
        /// Todos los items que se graban juntos en una misma anotación comparten el mismo IdAnotacion.
        /// </summary>
        public int IdAnotacion { get; set; }

        /// <summary>
        /// Número de pedido (PEDENCA.NroPed)
        /// </summary>
        public int NroPed { get; set; }

        /// <summary>
        /// Código interno del producto (PEDDETA.CodiInt)
        /// </summary>
        public int CodiInt { get; set; }

        /// <summary>
        /// Código externo del producto (MASTER.Codigo / CODEXT en VB6)
        /// </summary>
        [StringLength(24)]
        public string? CodiExt { get; set; }

        /// <summary>
        /// Descripción del producto (PEDDETA.Descrip / DESCRIT0 en VB6)
        /// </summary>
        [StringLength(384)]
        public string? Descrip { get; set; }

        /// <summary>
        /// Cantidad pedida (PEDDETA.CanPed)
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CanPed { get; set; }

        /// <summary>
        /// Cantidad entregada (PEDDETA.CantEnt)
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantEnt { get; set; }

        /// <summary>
        /// Observaciones particulares del item
        /// </summary>
        [StringLength(512)]
        public string? ObservaItem { get; set; }

        /// <summary>
        /// Observaciones generales del pedido (aplican a toda la anotación)
        /// </summary>
        [StringLength(1024)]
        public string? ObservaGral { get; set; }

        /// <summary>
        /// Fecha del cambio que originó la anotación
        /// </summary>
        public DateTime? FechaCambio { get; set; }

        /// <summary>
        /// Nombre del usuario activo que originó la anotación
        /// </summary>
        [StringLength(64)]
        public string? Usuario { get; set; }

        /// <summary>
        /// Código de empresa. Por convención del proyecto siempre es 0.
        /// </summary>
        public short? CodiEmpr { get; set; }

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
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar.
        /// Mapea a usuarios.idflexoft para compatibilidad con DER_ACCE (VB6).
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
-- =============================================================================
-- Script SQL para crear la tabla ANOTAPED en SQL Server (ejecutar manualmente)
-- =============================================================================

CREATE TABLE ANOTAPED (
    Id           INT IDENTITY(1,1) PRIMARY KEY,
    IdAnotacion  INT NOT NULL,
    NroPed       INT NOT NULL,
    CodiInt      INT NOT NULL,
    CodiExt      NVARCHAR(24) NULL,
    Descrip      NVARCHAR(384) NULL,
    CanPed       DECIMAL(18,4) NULL,
    CantEnt      DECIMAL(18,4) NULL,
    ObservaItem  NVARCHAR(512) NULL,
    ObservaGral  NVARCHAR(1024) NULL,
    FechaCambio  DATETIME NULL,
    Usuario      NVARCHAR(64) NULL,
    CodiEmpr     SMALLINT NULL DEFAULT 0,
    AddRecord    DATETIME NULL,
    LastUpdate   DATETIME NULL,
    CodiEmprNet  INT NULL DEFAULT 0,
    NumUsuar     INT NULL
);

CREATE INDEX IX_ANOTAPED_NroPed ON ANOTAPED (NroPed);
CREATE INDEX IX_ANOTAPED_NroPed_CodiInt ON ANOTAPED (NroPed, CodiInt);
CREATE INDEX IX_ANOTAPED_IdAnotacion ON ANOTAPED (IdAnotacion);
*/
