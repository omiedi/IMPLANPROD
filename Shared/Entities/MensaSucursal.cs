using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad MENSA_SUCURSAL - Tabla de mensajes de sucursal para órdenes de compra
    /// </summary>
    [Table("MENSA_SUCURSAL")]
    public class MensaSucursal : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador interno autoincremental (IDENTITY) y clave primaria
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Código de sucursal
        /// </summary>
        [Column("SUCURSAL")]
        public int? Sucursal { get; set; }

        /// <summary>
        /// Orden de compra de sucursal
        /// </summary>
        [Column("O_COMPRA_SUC")]
        public int? OCompraSuc { get; set; }

        /// <summary>
        /// Usuario que creó el mensaje
        /// </summary>
        [Column("USUARIO")]
        public int? Usuario { get; set; }

        /// <summary>
        /// Estado del registro
        /// </summary>
        [Column("STATUS")]
        public short? Status { get; set; }

        /// <summary>
        /// Mensaje de la sucursal
        /// </summary>
        [Column("MENSAJE")]
        [MaxLength(1024)]
        public string? Mensaje { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; } = 0;

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
INSTRUCCIONES SQL PARA CREAR LA TABLA:
Ejecutar en SQL Server Management Studio para crear la tabla MENSA_SUCURSAL:

-- CREAR TABLA MENSA_SUCURSAL
CREATE TABLE [dbo].[MENSA_SUCURSAL] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [SUCURSAL] INT NULL,
    [O_COMPRA_SUC] INT NULL,
    [USUARIO] INT NULL,
    [STATUS] SMALLINT NULL,
    [MENSAJE] NVARCHAR(1024) NULL,
    [CodiEmpr] INT NULL DEFAULT 0,
    [AddRecord] DATETIME2 NOT NULL DEFAULT GETDATE(),
    [LastUpdate] DATETIME2 NOT NULL DEFAULT GETDATE(),
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,
    CONSTRAINT PK_MENSA_SUCURSAL_Id PRIMARY KEY (Id)
);

-- ÍNDICES PARA MEJORAR RENDIMIENTO
-- Índice para filtrar por orden de compra y sucursal (usado en WHERE del modal)
CREATE NONCLUSTERED INDEX IX_MENSA_SUCURSAL_OCOMPRA_SUCURSAL ON [dbo].[MENSA_SUCURSAL] ([O_COMPRA_SUC], [SUCURSAL]);

-- NOTAS IMPORTANTES:
-- 1. La tabla se crea con un campo Id IDENTITY como clave primaria
-- 2. Los campos de texto usan NVARCHAR para soportar caracteres Unicode
-- 3. CodiEmpr tiene valor por defecto 0 según las reglas del sistema
-- 4. Los campos de auditoría (AddRecord, LastUpdate, CodiEmprNet, NumUsuar) se agregan automáticamente
-- 5. El índice compuesto optimiza las consultas para mostrar mensajes en el modal
-- 6. SUCURSAL es INT según el código VB6 (CREACAMPO tipo 2 = int)
-- 7. MENSAJE es NVARCHAR(1024) según el código VB6 (CREACAMPO longitud 1024)
*/
