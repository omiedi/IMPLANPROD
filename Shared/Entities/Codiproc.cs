using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa los Códigos de Proceso de Inspección (CODIPROC)
    /// Almacena los procedimientos de inspección con sus parámetros de calidad,
    /// frecuencias, puntos de detención y trazabilidad de quién preparó/aprobó el proceso
    /// 
    /// ORIGEN: Migrado desde archivo plano VB6/ISAM - CODIPROC.DAT (512 bytes por registro)
    /// FORM VB6 de referencia: PLAINSP07.frm
    /// </summary>
    [Table("CODIPROC")]
    public class Codiproc : ITimestampEntity, IAuditableEntity
    {
        // ===================================================================
        // CLAVE PRIMARIA (AGREGADA PARA SQL SERVER)
        // ===================================================================

        /// <summary>
        /// Identificador único auto-incremental (campo agregado para SQL Server)
        /// No existía en el archivo plano original
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("Id")]
        public int Id { get; set; }

        // ===================================================================
        // CAMPOS ORIGINALES DEL ARCHIVO PLANO (512 bytes)
        // ===================================================================

        /// <summary>
        /// Código de empresa
        /// No proviene del archivo plano; siempre 0 según reglas globales del sistema
        /// </summary>
        [Column("CodiEmpr")]
        public int CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Código interno del artículo/producto
        /// Posición en archivo plano: bytes 1-2 (CVI → Int16 little-endian)
        /// Leído con la función CVI de VB6: X = 256 * Asc(byte2) + Asc(byte1)
        /// Equivalente en C# a: BitConverter.ToInt16(bytes, offset)
        /// Relacionado con MASTER.CodInte
        /// Conversión: Int16 (archivo plano) → int (SQL/C#)
        /// </summary>
        [Column("Cod_Inte")]
        public int? Cod_Inte { get; set; }

        /// <summary>
        /// Código de proceso de inspección (clave de negocio)
        /// Posición en archivo plano: bytes 3-4 (smallint)
        /// Referenciado como Ci% en VB6 / Label28 en formulario PLAINSP07
        /// Índice 1 del archivo CODIPROC: RECFILT$("CODIPROC", 1, MKI$(Ci%))
        /// Conversión: smallint (archivo plano) → int (SQL/C#)
        /// </summary>
        [Column("CodProc")]
        public int? CodProc { get; set; }

        /// <summary>
        /// Nivel de Calidad Aceptable (AQL - Acceptable Quality Level)
        /// Posición en archivo plano: bytes 17-18 (CVI → smallint)
        /// Corresponde a Text1 en formulario PLAINSP07
        /// Conversión: smallint (archivo plano) → int (SQL/C#)
        /// </summary>
        [Column("AQL")]
        public int? AQL { get; set; }

        /// <summary>
        /// Nivel porcentual de inspección
        /// Posición en archivo plano: bytes 19-20 (CVI → smallint)
        /// Corresponde a Text3 en formulario PLAINSP07
        /// Conversión: smallint (archivo plano) → int (SQL/C#)
        /// </summary>
        [Column("NivelPct")]
        public int? NivelPct { get; set; }

        /// <summary>
        /// Frecuencia de Inspección
        /// Posición en archivo plano: bytes 21-52 (char 32 bytes)
        /// Corresponde a Text6 en formulario PLAINSP07
        /// Conversión: char(32) → nvarchar(32)
        /// </summary>
        [Column("FrecInspeccion")]
        [MaxLength(32)]
        public string? FrecInspeccion { get; set; }

        /// <summary>
        /// Puntos de Detención Obligatoria
        /// Posición en archivo plano: bytes 53-100 (char 48 bytes)
        /// Corresponde a Text7 en formulario PLAINSP07
        /// Conversión: char(48) → nvarchar(48)
        /// </summary>
        [Column("PuntoDetencion")]
        [MaxLength(48)]
        public string? PuntoDetencion { get; set; }

        /// <summary>
        /// Nombre del usuario que preparó el proceso de inspección
        /// Posición en archivo plano: bytes 101-124 (char 24 bytes)
        /// Corresponde a Text4(0) en formulario PLAINSP07
        /// Conversión: char(24) → nvarchar(24)
        /// </summary>
        [Column("NombrePreparo")]
        [MaxLength(24)]
        public string? NombrePreparo { get; set; }

        /// <summary>
        /// Fecha en que se preparó el proceso de inspección
        /// Posición en archivo plano: bytes 125-126 (CVI → fecha numérica VB6)
        /// Corresponde a Text5(0) en formulario PLAINSP07 (función FECHATEX$)
        /// Conversión: fecha numérica VB6 → smalldatetime (SQL) → DateTime? (C#)
        /// </summary>
        [Column("FechaPreparo")]
        public DateTime? FechaPreparo { get; set; }

        /// <summary>
        /// Nombre del usuario que aprobó el proceso de inspección
        /// Posición en archivo plano: bytes 127-150 (char 24 bytes)
        /// Corresponde a Text4(1) en formulario PLAINSP07
        /// Conversión: char(24) → nvarchar(24)
        /// </summary>
        [Column("NombreAprobo")]
        [MaxLength(24)]
        public string? NombreAprobo { get; set; }

        /// <summary>
        /// Fecha en que se aprobó el proceso de inspección
        /// Posición en archivo plano: bytes 151-152 (CVI → fecha numérica VB6)
        /// Corresponde a Text5(1) en formulario PLAINSP07 (función FECHATEX$)
        /// Conversión: fecha numérica VB6 → smalldatetime (SQL) → DateTime? (C#)
        /// </summary>
        [Column("FechaAprobo")]
        public DateTime? FechaAprobo { get; set; }

        /// <summary>
        /// Observaciones del proceso de inspección
        /// Posición en archivo plano: bytes 161-512 (char 352 bytes)
        /// Corresponde a Text2 en formulario PLAINSP07
        /// Conversión: char(352) → nvarchar(512) (ampliado para mayor capacidad)
        /// </summary>
        [Column("Observaciones")]
        [MaxLength(512)]
        public string? Observaciones { get; set; }

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
        /// Número de usuario que realizó la última modificación
        /// Corresponde a usuarios.idflexoft (usado en validación de derechos DER_ACCE)
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
===============================================================================
SCRIPT SQL - CREACIÓN DE TABLA CODIPROC
===============================================================================
ORIGEN: Migrado desde archivo plano VB6/ISAM CODIPROC.DAT (512 bytes/registro)
FORM VB6: PLAINSP07.frm - Sub CARENCAPLA()

NOTAS TÉCNICAS DEL ARCHIVO PLANO ORIGINAL:
- Tamaño de registro: 512 bytes
- Bytes 1-2:   Cod_Inte (CVI → Int16, código interno artículo ref MASTER.CodInte)
- Bytes 3-4:   CodProc  (smallint, clave búsqueda índice 1)
- Bytes 5-16:  (no mapeados en el formulario)
- Bytes 17-18: AQL - Nivel de Calidad Aceptable (CVI → smallint)
- Bytes 19-20: NivelPct - Nivel % (CVI → smallint)
- Bytes 21-52: FrecInspeccion - Frecuencia de Inspección (char 32)
- Bytes 53-100: PuntoDetencion - Puntos de Detención Obligatoria (char 48)
- Bytes 101-124: NombrePreparo - Nombre usuario que preparó (char 24)
- Bytes 125-126: FechaPreparo - Fecha preparó (fecha numérica VB6, CVI)
- Bytes 127-150: NombreAprobo - Nombre usuario que aprobó (char 24)
- Bytes 151-152: FechaAprobo - Fecha aprobó (fecha numérica VB6, CVI)
- Bytes 153-160: (no mapeados en el formulario)
- Bytes 161-512: Observaciones (char 352)

CONVERSIONES APLICADAS:
- smallint (archivo plano) → int (SQL/C#) para CodProc, AQL, NivelPct
- char(N) → nvarchar(N) para campos de texto
- fecha numérica VB6 → smalldatetime (SQL) → DateTime? (C#) para fechas
- CodiEmpr: smallint → int con valor por defecto 0
- Observaciones: char(352) → nvarchar(512) (ampliado)
===============================================================================

-- ============================================================
-- SCRIPT DE CREACIÓN: Ejecutar si la tabla NO existe
-- ============================================================
USE [IMPLANPROD]
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CODIPROC]') AND type = N'U')
BEGIN

    CREATE TABLE [dbo].[CODIPROC] (
        -- Clave primaria (agregada para SQL Server)
        [Id]              [int]           IDENTITY(1,1)   NOT NULL,

        -- Campos originales del archivo plano
        [CodiEmpr]        [int]           NOT NULL         DEFAULT 0,   -- Código empresa (siempre 0, no viene del .DAT)
        [Cod_Inte]        [int]           NULL,                          -- Código interno artículo (CVI pos 1-2, ref MASTER.CodInte)
        [CodProc]         [int]           NULL,                          -- Código de proceso (clave de negocio)
        [AQL]             [int]           NULL,                          -- Nivel de Calidad Aceptable
        [NivelPct]        [int]           NULL,                          -- Nivel porcentual de inspección
        [FrecInspeccion]  [nvarchar](32)  NULL,                          -- Frecuencia de Inspección
        [PuntoDetencion]  [nvarchar](48)  NULL,                          -- Puntos de Detención Obligatoria
        [NombrePreparo]   [nvarchar](24)  NULL,                          -- Nombre usuario que preparó
        [FechaPreparo]    [smalldatetime] NULL,                          -- Fecha en que se preparó
        [NombreAprobo]    [nvarchar](24)  NULL,                          -- Nombre usuario que aprobó
        [FechaAprobo]     [smalldatetime] NULL,                          -- Fecha en que se aprobó
        [Observaciones]   [nvarchar](512) NULL,                          -- Observaciones (ampliado a 512)

        -- Campos de auditoría
        [AddRecord]       [datetime]      NULL,                          -- Fecha de creación del registro
        [LastUpdate]      [datetime]      NULL,                          -- Fecha última modificación
        [CodiEmprNet]     [int]           NULL,                          -- Código empresa (auditoría)
        [NumUsuar]        [int]           NULL,                          -- Usuario que realizó la operación

        CONSTRAINT [PK_CODIPROC] PRIMARY KEY CLUSTERED ([Id] ASC)
    )

    -- Índice por CodProc (clave de negocio principal)
    CREATE NONCLUSTERED INDEX [IX_CODIPROC_CodProc]
        ON [dbo].[CODIPROC] ([CodProc] ASC)

    -- Índice por Cod_Inte (código interno artículo, para búsquedas cruzadas con MASTER)
    CREATE NONCLUSTERED INDEX [IX_CODIPROC_Cod_Inte]
        ON [dbo].[CODIPROC] ([Cod_Inte] ASC)

    -- Índice por CodiEmpr + CodProc (combinado, para búsquedas filtradas por empresa)
    CREATE NONCLUSTERED INDEX [IX_CODIPROC_CodiEmpr_CodProc]
        ON [dbo].[CODIPROC] ([CodiEmpr] ASC, [CodProc] ASC)

    PRINT '✅ Tabla CODIPROC creada exitosamente'
    PRINT '   Índices creados: IX_CODIPROC_CodProc, IX_CODIPROC_CodiEmpr_CodProc'
END
ELSE
BEGIN
    PRINT 'ℹ️ La tabla CODIPROC ya existe'
END
GO
//**** agregar este campo ************
ALTER TABLE CODIPROC ADD Cod_Inte int NULL


PRINT '============================================='
PRINT 'Script completado exitosamente'
PRINT 'Tabla: CODIPROC'
PRINT 'Descripción: Códigos de Proceso de Inspección'
PRINT 'Origen: Migrado desde archivo plano VB6 CODIPROC.DAT'
PRINT '============================================='
GO

===============================================================================
NOTAS DE MAPEO Y CONVERSIONES:
===============================================================================

CONVERSIONES APLICADAS:
- smallint (archivo plano) → int (SQL/C#) para CodProc, AQL, NivelPct
- char(N) → nvarchar(N) para FrecInspeccion, PuntoDetencion, NombrePreparo, NombreAprobo
- fecha numérica VB6 (CVI) → smalldatetime para FechaPreparo, FechaAprobo
- Observaciones: char(352) → nvarchar(512) (ampliado para mayor capacidad)
- CodiEmpr: smallint → int con valor por defecto 0

CAMPOS NO MAPEADOS DEL ARCHIVO PLANO ORIGINAL:
- Bytes 5-16  (12 bytes): No aparecen en el formulario PLAINSP07 - posiblemente reservados
- Bytes 153-160 (8 bytes): No aparecen en el formulario PLAINSP07 - posiblemente reservados

===============================================================================
*/
