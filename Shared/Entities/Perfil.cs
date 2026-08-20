using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    public class Perfil : ITimestampEntity, IAuditableEntity
    {
        public int Id { get; set; }

        [Required(ErrorMessage = "El nombre del perfil es obligatorio")]
        [StringLength(50, ErrorMessage = "El nombre del perfil no puede exceder los 50 caracteres")]
        public string Nombre { get; set; }

        public string Descripcion { get; set; }

        public List<Permiso> Permisos { get; set; } = new List<Permiso>();

        public List<Usuario> Usuarios { get; set; } = new List<Usuario>();

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
    IF NOT EXISTS (SELECT 1 FROM dbo.Perfiles WHERE UPPER(LTRIM(RTRIM(Nombre))) = 'SUPERADMIN')
    BEGIN
        INSERT INTO dbo.Perfiles
        (
            Nombre,
            Descripcion,
            AddRecord,
            LastUpdate,
            CodiEmprNet,
            NumUsuar
        )
        VALUES
        (
            'SUPERADMIN',
            'Acceso total. Unico perfil que puede asignar/revocar SUPERADMIN.',
            SYSDATETIME(),
            SYSDATETIME(),
            0,
            0
        );
    END;

 IF NOT EXISTS (SELECT 1 FROM dbo.Perfiles WHERE UPPER(LTRIM(RTRIM(Nombre))) = 'ADMINISTRADOR')
    BEGIN
        INSERT INTO dbo.Perfiles (Nombre, Descripcion, AddRecord, LastUpdate, CodiEmprNet, NumUsuar)
        VALUES ('ADMINISTRADOR', 'Acceso a todo menos configuracion de seguridad.', SYSDATETIME(), SYSDATETIME(), 0, 0);
    END;

    IF NOT EXISTS (SELECT 1 FROM dbo.Perfiles WHERE UPPER(LTRIM(RTRIM(Nombre))) = 'USUARIO')
    BEGIN
        INSERT INTO dbo.Perfiles (Nombre, Descripcion, AddRecord, LastUpdate, CodiEmprNet, NumUsuar)
        VALUES ('USUARIO', 'Acceso a rutas explicitamente asignadas.', SYSDATETIME(), SYSDATETIME(), 0, 0);
    END;

 */
