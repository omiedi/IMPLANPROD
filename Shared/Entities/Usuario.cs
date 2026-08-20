using System;
using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad Usuario con sistema de timestamps automáticos
    /// Implementa ITimestampEntity e IAuditableEntity para auditoría automática
    /// </summary>
    public class Usuario : ITimestampEntity, IAuditableEntity
    {
        public int Id { get; set; }

        [Required(ErrorMessage = "El nombre de usuario es obligatorio")]
        [StringLength(50, ErrorMessage = "El nombre de usuario no puede exceder los 50 caracteres")]
        public required string NombreUsuario { get; set; }

        [Required(ErrorMessage = "El correo electrónico es obligatorio")]
        [EmailAddress(ErrorMessage = "El formato del correo electrónico no es válido")]
        public required string Email { get; set; }

        [Required(ErrorMessage = "La contraseña es obligatoria")]
        [StringLength(100, MinimumLength = 6, ErrorMessage = "La contraseña debe tener entre 6 y 100 caracteres")]
        public required string Password { get; set; }

        // Campos para autenticación
        public string? RefreshToken { get; set; }
        public DateTime? RefreshTokenExpiryTime { get; set; }
        public DateTime UltimoAcceso { get; set; }
        public int IntentosLogin { get; set; }
        public bool Bloqueado { get; set; }

        public DateTime FechaRegistro { get; set; }
        public bool Activo { get; set; }
        public string? FotoUrl { get; set; }

        // Referencias geográficas
        [Required(ErrorMessage = "El país es obligatorio")]
        public int? CountryId { get; set; }
        public Country? Country { get; set; }

        [Required(ErrorMessage = "La provincia/estado es obligatorio")]
        public int? StateId { get; set; }
        public State? State { get; set; }

        [Required(ErrorMessage = "La ciudad es obligatoria")]
        public int? CityId { get; set; }
        public City? City { get; set; }

        [Required(ErrorMessage = "El perfil es obligatorio")]
        public int? PerfilId { get; set; }
        public Perfil? Perfil { get; set; }


        [Required(ErrorMessage = "La empresa es obligatoria")]
        public int EmpresaId { get; set; }
        public Empresa? Empresa { get; set; }

        [Required(ErrorMessage = "La dependencia es obligatoria")]
        public int DependenciaId { get; set; }
        public Dependencia? Dependencia { get; set; }

        [Required(ErrorMessage = "El sector es obligatorio")]
        public int SectorId { get; set; }
        public Sector? Sector { get; set; }

        public int IdFlexoft { get; set; }

        // Campos de timestamp automáticos - implementación de ITimestampEntity
        /// <summary>
        /// Fecha de creación del registro (se establece automáticamente al crear)
        /// </summary>
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro (se actualiza automáticamente al modificar)
        /// </summary>
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// </summary>
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// </summary>
        public int? NumUsuar { get; set; }
    }
}
/*
-- Instrucciones SQL para actualizar la tabla Usuarios:

-- 1. Agregar campo IdFlexoft (si no existe)
ALTER TABLE Usuarios
ADD IdFlexoft INT;

-- 2. Agregar campos de timestamp automáticos
ALTER TABLE Usuarios
ADD [AddRecord] datetime2 NOT NULL DEFAULT GETDATE(),
    [LastUpdate] datetime2 NOT NULL DEFAULT GETDATE();

-- 3. Crear índice para optimizar consultas por fecha de creación
CREATE INDEX IX_Usuarios_AddRecord ON Usuarios (AddRecord);

-- 4. Crear índice para optimizar consultas por fecha de actualización
CREATE INDEX IX_Usuarios_LastUpdate ON Usuarios (LastUpdate);

-- 5. PARA CREAR QUE EL USUARIO POR EJEMPLO VICTOR TENGA UN PERFIL DE SUPERADMIN, PRIMERO SE DEBE CREAR EL PERFIL SUPERADMIN **** update  Usuarios set PerfilId = 3 where id = 6 ***
(EN ESTE CASO EL ID 6 ES EL REGISTRO DE VICTOR Y EL ID 3 ES EL REGISTRO DEL PERFIL SUPERADMIN, ESTO ES SOLO UN EJEMPLO, LOS ID PUEDEN VARIAR SEGÚN LA BASE DE DATOS)



*/
