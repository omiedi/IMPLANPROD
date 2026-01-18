using DocumentFormat.OpenXml.Office2010.Excel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad para almacenar imágenes asociadas a modelos (BADAMODE)
    /// Tabla: GALIMMOD
    /// </summary>
    [Table("GALIMMOD")]
    public class GalimMod : IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public int CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Código del modelo (FK a BADAMODE)
        /// </summary>
        [MaxLength(16)]
        public string Cod_Mod { get; set; } = null!;

        /// <summary>
        /// Número de orden de la imagen (0, 1, 2, ...)
        /// </summary>
        public short Norit_Imag { get; set; }

        /// <summary>
        /// Nombre del archivo de imagen
        /// </summary>
        [MaxLength(255)]
        public string? Nom_Imag { get; set; }

        /// <summary>
        /// Ruta relativa de la imagen en el servidor
        /// (Se almacena la ruta en lugar del blob para mejor rendimiento)
        /// </summary>
        [MaxLength(500)]
        public string? RutaImagen { get; set; }

        /// <summary>
        /// Descripción u observaciones de la imagen
        /// </summary>
        [MaxLength(500)]
        public string? Descripcion { get; set; }

        // Campos de auditoría automática (IAuditableEntity)
        /// <summary>
        /// Fecha de creación del registro (asignada automáticamente por AuditInterceptor)
        /// </summary>
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última modificación (actualizada automáticamente por AuditInterceptor)
        /// </summary>
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa del usuario (asignado automáticamente por AuditInterceptor)
        /// </summary>
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// ID del usuario que creó/modificó (asignado automáticamente por AuditInterceptor)
        /// </summary>
        public int? NumUsuar { get; set; }

        // Nota: No se define relación de navegación con Badamode porque Cod_Mod es string
        // y no hay una FK real en la base de datos. La relación es lógica, no física.
    }
}
/*-- Crear tabla GALIMMOD
CREATE TABLE [dbo].[GALIMMOD] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [CodiEmpr] INT NOT NULL DEFAULT 0,
    [Cod_Mod] NVARCHAR(16) NOT NULL,
    [Norit_Imag] SMALLINT NOT NULL,
    [Nom_Imag] NVARCHAR(255) NULL,
    [RutaImagen] NVARCHAR(500) NULL,
    [Descripcion] NVARCHAR(500) NULL,
    [AddRecord] DATETIME NULL,
    [LastUpdate] DATETIME NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,
    
    CONSTRAINT [PK_GALIMMOD] PRIMARY KEY CLUSTERED ([Id] ASC)
);

-- Índice para búsquedas por código de modelo
CREATE NONCLUSTERED INDEX [IX_GALIMMOD_CodMod] 
    ON [dbo].[GALIMMOD] ([Cod_Mod] ASC);

-- Índice único para evitar duplicados de orden por modelo
CREATE UNIQUE NONCLUSTERED INDEX [IX_GALIMMOD_CodMod_Norit] 
    ON [dbo].[GALIMMOD] ([Cod_Mod] ASC, [Norit_Imag] ASC);

-- Índice para búsquedas por empresa
CREATE NONCLUSTERED INDEX [IX_GALIMMOD_CodiEmpr] 
    ON [dbo].[GALIMMOD] ([CodiEmpr] ASC);

GO-- Crear tabla GALIMMOD
CREATE TABLE [dbo].[GALIMMOD] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [CodiEmpr] INT NOT NULL DEFAULT 0,
    [Cod_Mod] NVARCHAR(16) NOT NULL,
    [Norit_Imag] SMALLINT NOT NULL,
    [Nom_Imag] NVARCHAR(255) NULL,
    [RutaImagen] NVARCHAR(500) NULL,
    [Descripcion] NVARCHAR(500) NULL,
    [AddRecord] DATETIME NULL,
    [LastUpdate] DATETIME NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,
    
    CONSTRAINT [PK_GALIMMOD] PRIMARY KEY CLUSTERED ([Id] ASC)
);

-- Índice para búsquedas por código de modelo
CREATE NONCLUSTERED INDEX [IX_GALIMMOD_CodMod] 
    ON [dbo].[GALIMMOD] ([Cod_Mod] ASC);

-- Índice único para evitar duplicados de orden por modelo
CREATE UNIQUE NONCLUSTERED INDEX [IX_GALIMMOD_CodMod_Norit] 
    ON [dbo].[GALIMMOD] ([Cod_Mod] ASC, [Norit_Imag] ASC);

-- Índice para búsquedas por empresa
CREATE NONCLUSTERED INDEX [IX_GALIMMOD_CodiEmpr] 
    ON [dbo].[GALIMMOD] ([CodiEmpr] ASC);

GO
*/