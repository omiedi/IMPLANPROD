using System;

namespace Shared.Entities
{
    /// <summary>
    /// Entidad para almacenar la definición y configuración de reportes personalizados.
    /// Permite guardar la estructura del reporte, condiciones y formato de columnas en formato JSON.
    /// </summary>
    public class ReportePersonalizado
    {
        /// <summary>
        /// Identificador único del reporte.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Nombre descriptivo del reporte, definido por el usuario.
        /// </summary>
        public string Nombre { get; set; }

        /// <summary>
        /// Nombre de usuario que creó el reporte (opcional).
        /// </summary>
        public string? Usuario { get; set; }

        /// <summary>
        /// Configuración completa del reporte en formato JSON (tablas, campos, relaciones, formato de columnas, etc).
        /// </summary>
        public string ConfiguracionJson { get; set; }

        /// <summary>
        /// Fecha de creación del reporte.
        /// </summary>
        public DateTime FechaCreacion { get; set; }

        /// <summary>
        /// Fecha de la última modificación del reporte.
        /// </summary>
        public DateTime FechaModificacion { get; set; }
    }
}
/*sQL Server
-- Crear tabla ReportePersonalizado (si no existe)

IF OBJECT_ID('dbo.ReportePersonalizado', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.ReportePersonalizado
    (
        Id                INT IDENTITY(1,1) NOT NULL,
        Nombre            NVARCHAR(200) NOT NULL,
        Usuario           NVARCHAR(100) NULL,
        ConfiguracionJson NVARCHAR(MAX) NOT NULL,
        FechaCreacion     DATETIME2(0) NOT NULL,
        FechaModificacion DATETIME2(0) NOT NULL,

        -- Campos estándar solicitados
        Codiempr          INT NOT NULL CONSTRAINT DF_ReportePersonalizado_Codiempr DEFAULT (0),
        AddRecord         DATETIME2(0) NULL,
        LastUpdate        DATETIME2(0) NULL,
        CodiEmprNet       INT NULL,
        NumUsuar          INT NULL,

        CONSTRAINT PK_ReportePersonalizado PRIMARY KEY CLUSTERED (Id)
    );
END
GO

-- Índices recomendados
IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_ReportePersonalizado_Nombre'
      AND object_id = OBJECT_ID('dbo.ReportePersonalizado')
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_ReportePersonalizado_Nombre
        ON dbo.ReportePersonalizado (Nombre);
END
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_ReportePersonalizado_FechaModificacion'
      AND object_id = OBJECT_ID('dbo.ReportePersonalizado')
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_ReportePersonalizado_FechaModificacion
        ON dbo.ReportePersonalizado (FechaModificacion DESC);
END
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_ReportePersonalizado_Codiempr_Nombre'
      AND object_id = OBJECT_ID('dbo.ReportePersonalizado')
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_ReportePersonalizado_Codiempr_Nombre
        ON dbo.ReportePersonalizado (Codiempr, Nombre);
END
GO
*/