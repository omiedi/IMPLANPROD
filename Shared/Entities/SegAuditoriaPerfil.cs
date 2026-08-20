using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("SEG_AUDITORIA_PERFIL")]
    public class SegAuditoriaPerfil : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public int UsuarioObjetivoId { get; set; }

        [ForeignKey(nameof(UsuarioObjetivoId))]
        public Usuario? UsuarioObjetivo { get; set; }

        public int? PerfilAnteriorId { get; set; }

        [ForeignKey(nameof(PerfilAnteriorId))]
        public Perfil? PerfilAnterior { get; set; }

        [Required]
        public int PerfilNuevoId { get; set; }

        [ForeignKey(nameof(PerfilNuevoId))]
        public Perfil? PerfilNuevo { get; set; }

        [Required]
        public int ModificadoPorUsuarioId { get; set; }

        [ForeignKey(nameof(ModificadoPorUsuarioId))]
        public Usuario? ModificadoPorUsuario { get; set; }

        public bool EsCambioSuperAdmin { get; set; }

        [MaxLength(400)]
        public string? Observacion { get; set; }

        public DateTime? FechaAccion { get; set; }

        [Column("Codiempr")]
        public int? Codiempr { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? LastUpdate { get; set; }

        public int? CodiEmprNet { get; set; }

        public int? NumUsuar { get; set; }
    }
}

/*
    IF OBJECT_ID('dbo.SEG_AUDITORIA_PERFIL', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.SEG_AUDITORIA_PERFIL
        (
            Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
            UsuarioObjetivoId INT NOT NULL,
            PerfilAnteriorId INT NULL,
            PerfilNuevoId INT NOT NULL,
            ModificadoPorUsuarioId INT NOT NULL,
            EsCambioSuperAdmin BIT NOT NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_EsCambioSuperAdmin DEFAULT(0),
            Observacion NVARCHAR(400) NULL,
            FechaAccion DATETIME2 NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_FechaAccion DEFAULT(SYSDATETIME()),
            Codiempr INT NOT NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_Codiempr DEFAULT(0),
            AddRecord DATETIME2 NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_AddRecord DEFAULT(SYSDATETIME()),
            LastUpdate DATETIME2 NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_LastUpdate DEFAULT(SYSDATETIME()),
            CodiEmprNet INT NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_CodiEmprNet DEFAULT(0),
            NumUsuar INT NULL CONSTRAINT DF_SEG_AUDITORIA_PERFIL_NumUsuar DEFAULT(0),
            CONSTRAINT FK_SEG_AUDITORIA_PERFIL_UsuarioObjetivo FOREIGN KEY (UsuarioObjetivoId) REFERENCES dbo.Usuarios(Id),
            CONSTRAINT FK_SEG_AUDITORIA_PERFIL_PerfilAnterior FOREIGN KEY (PerfilAnteriorId) REFERENCES dbo.Perfiles(Id),
            CONSTRAINT FK_SEG_AUDITORIA_PERFIL_PerfilNuevo FOREIGN KEY (PerfilNuevoId) REFERENCES dbo.Perfiles(Id),
            CONSTRAINT FK_SEG_AUDITORIA_PERFIL_ModificadoPor FOREIGN KEY (ModificadoPorUsuarioId) REFERENCES dbo.Usuarios(Id)
        );

        CREATE INDEX IX_SEG_AUDITORIA_PERFIL_Usuario_Fecha ON dbo.SEG_AUDITORIA_PERFIL(UsuarioObjetivoId, FechaAccion DESC);
    END;
*/
