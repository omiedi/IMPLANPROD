using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("SEG_USUARIO_PAGINA")]
    public class SegUsuarioPagina : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public int UsuarioId { get; set; }

        [ForeignKey(nameof(UsuarioId))]
        public Usuario? Usuario { get; set; }

        [Required]
        public int PaginaId { get; set; }

        [ForeignKey(nameof(PaginaId))]
        public SegPagina? Pagina { get; set; }

        public bool Permitido { get; set; } = true;

        [Column("Codiempr")]
        public int? Codiempr { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? LastUpdate { get; set; }

        public int? CodiEmprNet { get; set; }

        public int? NumUsuar { get; set; }
    }
}

/*
 * IF OBJECT_ID('dbo.SEG_USUARIO_PAGINA', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.SEG_USUARIO_PAGINA
        (
            Id          INT       IDENTITY(1,1) NOT NULL PRIMARY KEY,
            UsuarioId   INT       NOT NULL,
            PaginaId    INT       NOT NULL,
            Permitido   BIT       NOT NULL CONSTRAINT DF_SEG_USUARIO_PAGINA_Permitido   DEFAULT(1),
            Codiempr    INT       NOT NULL CONSTRAINT DF_SEG_USUARIO_PAGINA_Codiempr    DEFAULT(0),
            AddRecord   DATETIME2 NULL     CONSTRAINT DF_SEG_USUARIO_PAGINA_AddRecord   DEFAULT(SYSDATETIME()),
            LastUpdate  DATETIME2 NULL     CONSTRAINT DF_SEG_USUARIO_PAGINA_LastUpdate  DEFAULT(SYSDATETIME()),
            CodiEmprNet INT       NULL     CONSTRAINT DF_SEG_USUARIO_PAGINA_CodiEmprNet DEFAULT(0),
            NumUsuar    INT       NULL     CONSTRAINT DF_SEG_USUARIO_PAGINA_NumUsuar    DEFAULT(0),
            CONSTRAINT FK_SEG_USUARIO_PAGINA_Usuarios  FOREIGN KEY (UsuarioId) REFERENCES dbo.Usuarios(Id),
            CONSTRAINT FK_SEG_USUARIO_PAGINA_SEG_PAGINA FOREIGN KEY (PaginaId) REFERENCES dbo.SEG_PAGINA(Id)
        );

        CREATE UNIQUE INDEX UX_SEG_USUARIO_PAGINA_Usuario_Pagina ON dbo.SEG_USUARIO_PAGINA(UsuarioId, PaginaId);
    END;
*/
