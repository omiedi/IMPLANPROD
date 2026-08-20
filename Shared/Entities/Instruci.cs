using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad INSTRUCI - Instrucciones de Proceso de Inspección
    /// Migrado desde archivo plano VB6: INSTRUCI.DAT
    ///
    /// ESTRUCTURA DEL ARCHIVO PLANO (registro de longitud fija):
    /// VB6 usa REPLA(XX$, valor, posicion_base1, longitud)
    ///
    /// - Pos   1 (16 bytes): Codigo         → varchar(16)  left$(label37)
    /// - Pos  17 (48 bytes): Descripcion    → varchar(48)  AJUSTI$(Text3, 48)
    /// - Pos  65 ( 4 bytes): Revision       → varchar(4)   AJUSTI$(Text1, 4)
    /// - Pos  69 ( 2 bytes): Cod_Inte       → int          AJUSTI$(LABEL28, 4) len 2 (CVI)
    /// - Pos  71 (24 bytes): Preparo        → varchar(24)  AJUSTI$(Text4(0), 24)
    /// - Pos  95 ( 2 bytes): FechaPrep      → Int16 MKI$(FECHANUM(Text5(0))) → DateTime
    /// - Pos  97 (24 bytes): Aprobo         → varchar(24)  AJUSTI$(Text4(1), 24)
    /// - Pos 121 ( 2 bytes): FechaPro       → Int16 MKI$(FECHANUM(Text5(1))) → DateTime
    ///
    /// NOTAS DE CONVERSIÓN:
    /// - FECHANUM produce un Int16 con el mismo formato que FECHATEX$:
    ///   FEX = Día + 32*Mes + 416*(Año2d - 80)
    ///   Se lee con ConvertFechatexToDateTime (BitConverter.ToInt16 + decode)
    /// - AJUSTI$ rellena el string a la longitud indicada (padding/trim)
    /// - left$(label37): caption del label37 del formulario, usado como código identificador
    /// </summary>
    [Table("INSTRUCI")]
    public class Instruci : ITimestampEntity, IAuditableEntity
    {
        // ===================================================================
        // CLAVE PRIMARIA
        // ===================================================================

        /// <summary>
        /// Identificador único auto-incremental (no existe en el archivo plano,
        /// se genera al migrar a SQL)
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("Id")]
        public int Id { get; set; }

        // ===================================================================
        // CAMPOS DEL ARCHIVO PLANO - en orden de posición
        // ===================================================================

        /// <summary>
        /// Código de empresa
        /// No proviene del archivo plano; siempre 0 según reglas globales del sistema
        /// </summary>
        [Column("CodiEmpr")]
        public int CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Código identificador de la instrucción
        /// Posición en archivo plano: bytes 1-16 (VB6 pos 1, len 16)
        /// Origen VB6: left$(label37)
        /// </summary>
        [Column("Codigo")]
        [MaxLength(16)]
        public string? Codigo { get; set; }

        /// <summary>
        /// Descripción / texto de la instrucción de proceso
        /// Posición en archivo plano: bytes 17-64 (VB6 pos 17, len 48)
        /// Origen VB6: AJUSTI$(Text3, 48)
        /// </summary>
        [Column("Descripcion")]
        [MaxLength(48)]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Número o código de revisión (4 caracteres)
        /// Posición en archivo plano: bytes 65-68 (VB6 pos 65, len 4)
        /// Origen VB6: AJUSTI$(Text1, 4)
        /// </summary>
        [Column("Revision")]
        [MaxLength(4)]
        public string? Revision { get; set; }

        /// <summary>
        /// Código interno del artículo/producto al que aplica esta instrucción
        /// Posición en archivo plano: bytes 69-70 (VB6 pos 69, len 2)
        /// Leído con CVI de VB6 → BitConverter.ToInt16 (little-endian)
        /// Relacionado con MASTER.CodInte y CODIPROC.Cod_Inte
        /// Conversión: Int16 (archivo plano) → int (SQL/C#)
        /// </summary>
        [Column("Cod_Inte")]
        public int? Cod_Inte { get; set; }

        /// <summary>
        /// Nombre de quien preparó la instrucción
        /// Posición en archivo plano: bytes 71-94 (VB6 pos 71, len 24)
        /// Origen VB6: AJUSTI$(Text4(0), 24)
        /// </summary>
        [Column("Preparo")]
        [MaxLength(24)]
        public string? Preparo { get; set; }

        /// <summary>
        /// Fecha en que se preparó la instrucción
        /// Posición en archivo plano: bytes 95-96 (VB6 pos 95, len 2)
        /// Origen VB6: MKI$(FECHANUM(Text5(0))) → Int16 codificado con formato FECHATEX$
        /// Decodificación: FEX = Día + 32*Mes + 416*(Año2d-80)
        /// </summary>
        [Column("FechaPrep")]
        public DateTime? FechaPrep { get; set; }

        /// <summary>
        /// Nombre de quien aprobó la instrucción
        /// Posición en archivo plano: bytes 97-120 (VB6 pos 97, len 24)
        /// Origen VB6: AJUSTI$(Text4(1), 24)
        /// </summary>
        [Column("Aprobo")]
        [MaxLength(24)]
        public string? Aprobo { get; set; }

        /// <summary>
        /// Fecha en que se aprobó la instrucción
        /// Posición en archivo plano: bytes 121-122 (VB6 pos 121, len 2)
        /// Origen VB6: MKI$(FECHANUM(Text5(1))) → Int16 codificado con formato FECHATEX$
        /// Decodificación: FEX = Día + 32*Mes + 416*(Año2d-80)
        /// </summary>
        [Column("FechaPro")]
        public DateTime? FechaPro { get; set; }

        /// <summary>
        /// Observaciones / notas extendidas (campo adicional, no existe en el archivo plano VB6)
        /// Permite almacenar texto largo sin límite de longitud
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
        /// Número de usuario activo que creó/modificó el registro
        /// Se corresponde con Usuarios.IdFlexoft (numeración del sistema legacy)
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
============================================================
SCRIPT SQL - Crear tabla INSTRUCI en SQL Server
============================================================
Ejecutar UNA SOLA VEZ en la base de datos de destino.

CREATE TABLE [dbo].[INSTRUCI] (
    [Id]            [int]           IDENTITY(1,1)   NOT NULL,   -- PK auto-incremental
    [CodiEmpr]      [int]           NOT NULL         DEFAULT 0,  -- Código empresa (siempre 0)
    [Codigo]        [nvarchar](16)  NULL,                        -- Código identificador (pos 1-16)
    [Descripcion]   [nvarchar](48)  NULL,                        -- Descripción/Instrucción (pos 17-64)
    [Revision]      [nvarchar](4)   NULL,                        -- Número/código de revisión (pos 65-68)
    [Cod_Inte]      [int]           NULL,                        -- Código interno artículo (pos 69-70, CVI, ref MASTER.CodInte)
    [Preparo]       [nvarchar](24)  NULL,                        -- Nombre preparador (pos 71-94)
    [FechaPrep]     [datetime]      NULL,                        -- Fecha preparación (pos 95-96, FECHATEX)
    [Aprobo]        [nvarchar](24)  NULL,                        -- Nombre aprobador (pos 97-120)
    [FechaPro]      [datetime]      NULL,                        -- Fecha aprobación (pos 121-122, FECHATEX)
    [Observaciones] [nvarchar](MAX) NULL,                        -- Campo extendido (no existe en VB6)
    [AddRecord]     [datetime]      NULL,                        -- Fecha creación registro
    [LastUpdate]    [datetime]      NULL,                        -- Fecha última modificación
    [CodiEmprNet]   [int]           NULL,                        -- Código empresa (auditoría)
    [NumUsuar]      [int]           NULL,                        -- Usuario activo (ref. Usuarios.IdFlexoft)
    CONSTRAINT [PK_INSTRUCI] PRIMARY KEY CLUSTERED ([Id] ASC)
);
GO

-- Índice por Cod_Inte para joins con MASTER y CODIPROC
CREATE NONCLUSTERED INDEX [IX_INSTRUCI_Cod_Inte]
    ON [dbo].[INSTRUCI] ([Cod_Inte] ASC)
    WHERE [Cod_Inte] IS NOT NULL;
GO

-- Índice por Codigo para búsquedas frecuentes
CREATE NONCLUSTERED INDEX [IX_INSTRUCI_Codigo]
    ON [dbo].[INSTRUCI] ([Codigo] ASC)
    WHERE [Codigo] IS NOT NULL;
GO

-- Índice por CodiEmpr para compatibilidad con filtros multi-empresa
CREATE NONCLUSTERED INDEX [IX_INSTRUCI_CodiEmpr]
    ON [dbo].[INSTRUCI] ([CodiEmpr] ASC);
GO

============================================================
*/
