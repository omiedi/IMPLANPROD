using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa las Dimensiones Según Plano para Inspección (PLAINSP)
    /// Cada registro es un ítem de la lista de control dimensional que se verifica
    /// durante el proceso de inspección de un artículo específico.
    ///
    /// ORIGEN: Migrado desde archivo plano VB6/ISAM - PLAINSP.DAT (128 bytes por registro)
    /// FORM VB6 de referencia: PLAINSP07.frm - List1 (lista de ítems de inspección)
    /// RELACIÓN: Cod_Inte → CODIPROC.Cod_Inte (código interno del artículo)
    ///
    /// Cómo se cargaba en VB6 (Sub CARENCAPLA):
    ///   For U& = 1 To Len(RNC$) Step 4
    ///     RENOTA& = CVS(Mid$(RNC$, U&, 4))         ' nro de registro PLAINSP (CVS = MBF float)
    ///     Y$ = REGLEIDO$("PLAINSP", RENOTA&)         ' leer registro
    ///     Mid$(Y$, 3, 32)  → DimSegunPlano
    ///     Mid$(Y$, 35, 30) → Instrumento
    ///     Mid$(Y$, 65, 64) → Observaciones
    ///   Next U&
    /// </summary>
    [Table("PLAINSP")]
    public class Plainsp : ITimestampEntity, IAuditableEntity
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
        // CAMPOS ORIGINALES DEL ARCHIVO PLANO (128 bytes)
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
        /// Clave foránea lógica hacia CODIPROC.Cod_Inte
        /// Permite saber a qué proceso de inspección pertenece este ítem dimensional
        /// </summary>
        [Column("Cod_Inte")]
        public int? Cod_Inte { get; set; }

        /// <summary>
        /// Número de orden del ítem dentro de la lista de inspección
        /// Representa la posición secuencial del ítem para el mismo Cod_Inte
        /// En VB6 correspondía al índice U& dentro del loop sobre RNC$
        /// </summary>
        [Column("NroOrden")]
        public int? NroOrden { get; set; }

        /// <summary>
        /// Dimensión según plano que se debe verificar
        /// Posición en archivo plano: bytes 3-34 (char 32)
        /// Corresponde a Mid$(Y$, 3, 32) en VB6 / columna 1 del List1
        /// </summary>
        [Column("DimSegunPlano")]
        public string? DimSegunPlano { get; set; }

        /// <summary>
        /// Instrumento de medición a utilizar para verificar la dimensión
        /// Posición en archivo plano: bytes 35-64 (char 30)
        /// Corresponde a Mid$(Y$, 35, 30) en VB6 / columna 2 del List1
        /// </summary>
        [Column("Instrumento")]
        public string? Instrumento { get; set; }

        /// <summary>
        /// Observaciones adicionales sobre la inspección dimensional
        /// Posición en archivo plano: bytes 65-128 (char 64)
        /// Corresponde a Mid$(Y$, 65, 64) en VB6 / columna 3 del List1
        /// </summary>
        [Column("Observaciones")]
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
===============================================================================
SCRIPT SQL - CREACIÓN DE TABLA PLAINSP
===============================================================================
ORIGEN: Migrado desde archivo plano VB6/ISAM PLAINSP.DAT (128 bytes/registro)
FORM VB6: PLAINSP07.frm

NOTAS TÉCNICAS DEL ARCHIVO PLANO ORIGINAL:
- Tamaño de registro: 128 bytes
- Bytes  1-2:   Cod_Inte    (CVI → Int16, código interno artículo, ref CODIPROC.Cod_Inte)
- Bytes  3-34:  DimSegunPlano → Mid$(Y$, 3, 32)   (char 32)
- Bytes 35-64:  Instrumento   → Mid$(Y$, 35, 30)  (char 30)
- Bytes 65-128: Observaciones → Mid$(Y$, 65, 64)  (char 64)

CONDICIÓN DE SALTO EN MIGRACIÓN:
  Se omite el registro si DimSegunPlano y Instrumento están ambos vacíos
  (indica registro eliminado o ranura vacía en el archivo ISAM)

RELACIÓN:
  PLAINSP.Cod_Inte → CODIPROC.Cod_Inte  (N registros PLAINSP por 1 CODIPROC)

===============================================================================

USE [TuBaseDeDatos]
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PLAINSP]') AND type = N'U')
BEGIN

    CREATE TABLE [dbo].[PLAINSP] (
        -- Clave primaria (agregada para SQL Server)
        [Id]            [int]           IDENTITY(1,1)   NOT NULL,

        -- Campos originales del archivo plano
        [CodiEmpr]      [int]           NOT NULL         DEFAULT 0,    -- Código empresa (siempre 0, no viene del .DAT)
        [Cod_Inte]      [int]           NULL,                           -- Cód. interno artículo (CVI pos 1-2, ref CODIPROC)
        [NroOrden]      [int]           NULL,                           -- Nro de orden secuencial dentro del Cod_Inte
        [DimSegunPlano] [nvarchar](64)  NULL,                           -- Dimensión según plano (Mid$ pos 3, largo 32)
        [Instrumento]   [nvarchar](64)  NULL,                           -- Instrumento de medición (Mid$ pos 35, largo 30)
        [Observaciones] [nvarchar](128) NULL,                           -- Observaciones (Mid$ pos 65, largo 64)

        -- Campos de auditoría (ITimestampEntity + IAuditableEntity)
        [AddRecord]     [datetime]      NULL,                           -- Fecha de creación
        [LastUpdate]    [datetime]      NULL,                           -- Fecha de última modificación
        [CodiEmprNet]   [int]           NULL,                           -- Código empresa (auditoría)
        [NumUsuar]      [int]           NULL,                           -- Usuario que realizó la operación

        CONSTRAINT [PK_PLAINSP] PRIMARY KEY CLUSTERED ([Id] ASC)
    )

    -- Índice por Cod_Inte (FK lógica hacia CODIPROC, para búsquedas por artículo)
    CREATE NONCLUSTERED INDEX [IX_PLAINSP_Cod_Inte]
        ON [dbo].[PLAINSP] ([Cod_Inte] ASC)

    -- Índice compuesto Cod_Inte + NroOrden (para recuperar la lista ordenada por artículo)
    CREATE NONCLUSTERED INDEX [IX_PLAINSP_Cod_Inte_NroOrden]
        ON [dbo].[PLAINSP] ([Cod_Inte] ASC, [NroOrden] ASC)

    PRINT '✅ Tabla PLAINSP creada exitosamente'
END
ELSE
BEGIN
    PRINT 'ℹ️ La tabla PLAINSP ya existe'
END
GO

PRINT '============================================='
PRINT 'Script completado exitosamente'
PRINT '============================================='
*/
