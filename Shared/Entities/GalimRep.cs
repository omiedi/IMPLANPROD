using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad para almacenar la galería de imágenes de reparaciones
    /// Tabla: GALIMREP
    /// Implementa IAuditableEntity para auditoría automática
    /// </summary>
    /// 
    [Table("GalimRep")]
    public class GalimRep : IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public int CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Número de reparación (FK a REPARA)
        /// </summary>
        public int NumRepa { get; set; }

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

        // Relación con Repara
        [ForeignKey(nameof(NumRepa))]
        public Repara? Repara { get; set; }
    }
}

/*
-- Script SQL para crear la tabla GALIMREP en SQL Server

CREATE TABLE [dbo].[GALIMREP] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [CodiEmpr] INT NOT NULL,
    [NumRepa] INT NOT NULL,
    [Norit_Imag] SMALLINT NOT NULL,
    [Nom_Imag] NVARCHAR(255) NULL,
    [RutaImagen] NVARCHAR(500) NULL,
    [Descripcion] NVARCHAR(500) NULL,
    
    -- Campos de auditoría automática (asignados por AuditInterceptor)
    [AddRecord] DATETIME NULL,           -- Fecha de creación
    [LastUpdate] DATETIME NULL,          -- Fecha de última modificación
    [CodiEmprNet] INT NULL,              -- Código de empresa del usuario
    [NumUsuar] INT NULL,                 -- ID del usuario que creó/modificó
    
    CONSTRAINT [PK_GALIMREP] PRIMARY KEY CLUSTERED ([Id] ASC),
    
  
);

-- Índice para búsquedas por NumRepa (consulta principal)
CREATE NONCLUSTERED INDEX [IX_GALIMREP_NumRepa] 
    ON [dbo].[GALIMREP] ([NumRepa] ASC);

-- Índice único compuesto para evitar duplicados de orden de imagen por reparación
CREATE UNIQUE NONCLUSTERED INDEX [IX_GALIMREP_NumRepa_Norit] 
    ON [dbo].[GALIMREP] ([NumRepa] ASC, [Norit_Imag] ASC);

-- Índice para búsquedas por empresa
CREATE NONCLUSTERED INDEX [IX_GALIMREP_CodiEmpr] 
    ON [dbo].[GALIMREP] ([CodiEmpr] ASC);

GO
*/
