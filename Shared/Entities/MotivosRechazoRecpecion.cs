using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("MotivosRechazoRecepcion")]
    public class MotivosRechazoRecepcion : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Display(Name = "Código Motivo")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public int CodigoMotivo { get; set; }

        [Display(Name = "Descripción")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(128, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? Descripcion { get; set; }

        [Display(Name = "Estado")]
        public int Status { get; set; } = 0;

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
 -- =============================================
-- Script de creación de tabla MotivosRechazoRecepcion
-- Descripción: Tabla para almacenar los motivos de rechazo en recepciones
-- =============================================

-- 1. Verificar si la tabla existe y eliminarla (opcional, comentar si no se desea)
-- IF OBJECT_ID(N'[dbo].[MotivosRechazoRecepcion]', N'U') IS NOT NULL
--     DROP TABLE [dbo].[MotivosRechazoRecepcion];
-- GO

-- 2. Crear la tabla
CREATE TABLE [dbo].[MotivosRechazoRecepcion] (
    -- Clave primaria autoincremental
    [Id] INT IDENTITY(1,1) NOT NULL,
    
    -- Código del motivo
    [CodigoMotivo] INT NOT NULL,
    
    -- Descripción del motivo (128 caracteres)
    [Descripcion] NVARCHAR(128) NOT NULL,
    
    -- Estado (0 = Activo, -1 = Inactivo)
    [Status] INT NOT NULL DEFAULT 0,
    
    -- Constraint de clave primaria
    CONSTRAINT [PK_MotivosRechazoRecepcion] PRIMARY KEY CLUSTERED ([Id] ASC)
        WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, 
              ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY];
GO

-- 3. Crear índices para mejorar el rendimiento

-- Índice en CodigoMotivo (para búsquedas por código)
CREATE NONCLUSTERED INDEX [IX_MotivosRechazoRecepcion_CodigoMotivo] 
    ON [dbo].[MotivosRechazoRecepcion] ([CodigoMotivo] ASC);
GO

-- Índice en Status (para filtrar por estado activo/inactivo)
CREATE NONCLUSTERED INDEX [IX_MotivosRechazoRecepcion_Status] 
    ON [dbo].[MotivosRechazoRecepcion] ([Status] ASC);
GO

-- 4. Agregar constraint único para CodigoMotivo
ALTER TABLE [dbo].[MotivosRechazoRecepcion]
    ADD CONSTRAINT [UQ_MotivosRechazoRecepcion_CodigoMotivo] 
    UNIQUE ([CodigoMotivo]);
GO

-- 5. Insertar datos por defecto (ejemplos de motivos comunes)
INSERT INTO [dbo].[MotivosRechazoRecepcion] ([CodigoMotivo], [Descripcion], [Status])
VALUES 
    (1, 'Defecto de calidad', 1),
    (2, 'Producto dañado', 1),
    (3, 'No cumple especificaciones', 1),
    (4, 'Vencido o próximo a vencer', 1),
    (5, 'Error en el pedido', 1),
    (6, 'Embalaje defectuoso', 1),
    (7, 'Producto incorrecto', 1),
    (8, 'Cantidad incorrecta', 1),
    (9, 'Sin documentación requerida', 1),
    (10, 'Otros', 1);
GO

-- 6. Verificar la creación
SELECT 
    COLUMN_NAME,
    DATA_TYPE,
    CHARACTER_MAXIMUM_LENGTH,
    IS_NULLABLE,
    COLUMN_DEFAULT
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'MotivosRechazoRecepcion'
ORDER BY ORDINAL_POSITION;
GO

-- 7. Verificar los índices creados
SELECT 
    i.name AS IndexName,
    i.type_desc AS IndexType,
    c.name AS ColumnName
FROM sys.indexes i
INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
INNER JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id
WHERE i.object_id = OBJECT_ID(N'[dbo].[MotivosRechazoRecepcion]')
ORDER BY i.name, ic.key_ordinal;
GO

-- 8. Consulta de verificación con datos insertados
SELECT 
    Id,
    CodigoMotivo,
    Descripcion,
    CASE Status 
        WHEN 1 THEN 'Activo' 
        WHEN 0 THEN 'Inactivo'
        ELSE 'Baja' 
    END AS EstadoDescripcion,
    Status
FROM [dbo].[MotivosRechazoRecepcion]
ORDER BY CodigoMotivo;
GO
*/
 