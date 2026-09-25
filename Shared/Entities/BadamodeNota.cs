using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad BADAMODE_NOTA - Notas y observaciones de moldes
    /// Tabla que almacena las anotaciones y observaciones relacionadas con los moldes
    /// Implementa interfaces de auditoría para trazabilidad automática
    /// </summary>
    [Table("BADAMODE_NOTA")]
    public class BadamodeNota : ITimestampEntity, IAuditableEntity
    {
        // ===================================================================
        // CLAVE PRIMARIA
        // ===================================================================

        /// <summary>
        /// Identificador único autoincremental de la nota
        /// Clave primaria de la tabla
        /// </summary>
        [Key]
        [Column("Id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        // ===================================================================
        // CAMPOS DE DATOS
        // ===================================================================

        /// <summary>
        /// Código del molde asociado a esta nota
        /// Relación con la tabla BADAMODE
        /// </summary>
        [Column("Cod_Mod")]
        [StringLength(24)]
        public string? Cod_Mod { get; set; }

        /// <summary>
        /// Observaciones y anotaciones sobre el molde
        /// Campo de texto largo para almacenar información detallada
        /// </summary>
        [Column("Observaciones")]
        [StringLength(int.MaxValue)]
        public string? Observaciones { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// Se establece automáticamente cuando se crea la nota
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se actualiza automáticamente cada vez que se modifica la nota
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        // ===================================================================
        // CAMPOS DE AUDITORÍA - Implementación de IAuditableEntity
        // ===================================================================

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// Se actualiza automáticamente con la empresa del usuario actual
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// Se actualiza automáticamente con el ID del usuario actual
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
-- =============================================
-- Script: Creación de Tabla BADAMODE_NOTA
-- Descripción: Notas y observaciones de moldes
-- Fecha: 2026-01-08
-- =============================================

-- Crear la tabla BADAMODE_NOTA
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE_NOTA]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[BADAMODE_NOTA](
        -- Clave primaria autoincremental
        [Id] INT IDENTITY(1,1) NOT NULL,
        
        -- Campos de datos
        [Cod_Mod] NVARCHAR(24) NULL,
        [Observaciones] NVARCHAR(MAX) NULL,
        
        -- Campos de auditoría
        [AddRecord] DATETIME NULL,
        [LastUpdate] DATETIME NULL,
        [CodiEmprNet] INT NULL,
        [NumUsuar] INT NULL,
        
        -- Constraint de clave primaria
        CONSTRAINT [PK_BADAMODE_NOTA] PRIMARY KEY CLUSTERED ([Id] ASC)
    ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
    
    PRINT 'Tabla BADAMODE_NOTA creada exitosamente'
END
ELSE
BEGIN
    PRINT 'La tabla BADAMODE_NOTA ya existe'
END
GO

-- Crear índice en Cod_Mod para búsquedas rápidas por molde
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE_NOTA]') AND name = 'IX_BADAMODE_NOTA_Cod_Mod')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BADAMODE_NOTA_Cod_Mod]
    ON [dbo].[BADAMODE_NOTA] ([Cod_Mod])
    
    PRINT 'Índice IX_BADAMODE_NOTA_Cod_Mod creado'
END
ELSE
BEGIN
    PRINT 'El índice IX_BADAMODE_NOTA_Cod_Mod ya existe'
END
GO

-- Crear índice en AddRecord para ordenamiento por fecha
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BADAMODE_NOTA]') AND name = 'IX_BADAMODE_NOTA_AddRecord')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BADAMODE_NOTA_AddRecord]
    ON [dbo].[BADAMODE_NOTA] ([AddRecord] DESC)
    
    PRINT 'Índice IX_BADAMODE_NOTA_AddRecord creado'
END
ELSE
BEGIN
    PRINT 'El índice IX_BADAMODE_NOTA_AddRecord ya existe'
END
GO

-- Verificación final
SELECT 
    c.name AS 'Columna',
    t.name AS 'Tipo',
    c.max_length AS 'Longitud',
    c.is_nullable AS 'Permite NULL',
    c.is_identity AS 'Es IDENTITY'
FROM sys.columns c
INNER JOIN sys.types t ON c.user_type_id = t.user_type_id
WHERE c.object_id = OBJECT_ID(N'[dbo].[BADAMODE_NOTA]')
ORDER BY c.column_id

PRINT '============================================='
PRINT 'Script completado exitosamente'
PRINT 'Tabla BADAMODE_NOTA lista para usar'
PRINT '============================================='
GO


======================== ESTA INSTRUCCION ELIMINA EL CAMPO ROWGUID Y EL DE LAS NOTAS ANOTAMOD ===================================

DECLARE @tabla SYSNAME = 'badamode';

-- =========================
-- ELIMINAR CONSTRAINTS DEFAULT
-- =========================
DECLARE @sql NVARCHAR(MAX) = '';

SELECT @sql = @sql + '
ALTER TABLE ' + QUOTENAME(@tabla) +
' DROP CONSTRAINT ' + QUOTENAME(dc.name) + ';'
FROM sys.default_constraints dc
JOIN sys.columns c 
    ON dc.parent_object_id = c.object_id 
   AND dc.parent_column_id = c.column_id
WHERE dc.parent_object_id = OBJECT_ID(@tabla)
  AND c.name IN ('rowguid', 'anotamod');

IF (@sql <> '')
BEGIN
    PRINT 'Eliminando constraints DEFAULT...';
    EXEC sp_executesql @sql;
END

-- =========================
-- ELIMINAR COLUMNAS SI EXISTEN
-- =========================
IF COL_LENGTH(@tabla, 'rowguid') IS NOT NULL
BEGIN
    PRINT 'Eliminando columna rowguid';
    ALTER TABLE badamode DROP COLUMN rowguid;
END

IF COL_LENGTH(@tabla, 'anotamod') IS NOT NULL
BEGIN
    PRINT 'Eliminando columna anotamod';
    ALTER TABLE badamode DROP COLUMN anotamod;
END

====================================== LUEGO CRE DE NUEVO EL SCRIPT EN ARCON SIN LOS CAMPOS (ELIMINADOS) Y EJECUTE EL SCRIPT PARA INSERTAR LOS DATOS ===========

*/
// *** --- MIGRACION DE DATOS DESDE BADAMODE A BADAMODE_NOTA --- ***
// para migrar de la tabla BADAMODE a BADAMODE de arcon a aarcon2
/*INSERT INTO ARCON2.dbo.BADAMODE
(
    Cod_Mod,
    CodiEmpr,
    Descripcion,
    Ubic,
    CantPiezas,
    Frecuen,
    FechUlt_Con,
    FechPro_Con,
    Material,
    Desc_Mat,
    Cod_Inte,
    AnotaMod,
    ObservaMod,
    Foto1,
    Foto2,
    Foto3,
    Foto4,
    rowguid
)
SELECT
    Cod_Mod,
    CodiEmpr,
    Descripcion,
    Ubic,
    CantPiezas,
    Frecuen,
    FechUlt_Con,
    FechPro_Con,
    Material,
    Desc_Mat,
    Cod_Inte,
    AnotaMod,
    ObservaMod,
    Foto1,
    Foto2,
    Foto3,
    Foto4,
    NEWID()
FROM ARCON.dbo.BADAMODE;
*/
/*  ESTO ES LPARA CARGAR LAS NOTAS EN LA TABLA BADAMODE_NOTA  
 *  --1) Borrar el contenido de la tabla destino
DELETE FROM arcon2.dbo.BADAMODE_NOTA;


--2) Insertar los datos desde la tabla original
INSERT INTO arcon2.dbo.BADAMODE_NOTA (COD_MOD, OBSERVACIONES)
SELECT 
    COD_MOD,
    ANOTAMOD
FROM arcon.dbo.BADAMODE
WHERE ANOTAMOD IS NOT NULL
      AND ANOTAMOD <> '';
*/
