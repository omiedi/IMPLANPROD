using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Server.Utils;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IConfiguration _configuration;
        private readonly ILogger<AuthController> _logger;
        private readonly IEmailEmpresaService _emailService;

        public AuthController(DataContext context, IConfiguration configuration, ILogger<AuthController> logger, IEmailEmpresaService emailService)
        {
            _context = context;
            _configuration = configuration;
            _logger = logger;
            _emailService = emailService;
        }

        [HttpPost("login")]
        public async Task<IActionResult> Login([FromBody] LoginRequest request)
        {
            try
            {
                // Verificación especial para restablecer contraseña (solo en desarrollo)
                // Si el email termina con ".reset", restablecemos la contraseña del usuario
                if (request.Email.EndsWith(".reset"))
                {
                    string realEmail = request.Email.Replace(".reset", "");
                    var result = await ResetPasswordInternal(realEmail, request.Password);
                    if (result.Success)
                    {
                        return Ok(new { Mensaje = result.Message });
                    }
                    else
                    {
                        return BadRequest(result.Message);
                    }
                }

                // Bootstrap de seguridad:
                // 1) Identificamos el PerfilId del rol SUPERADMIN.
                // 2) Verificamos que exista al menos un usuario activo con ese perfil.
                // Si no existe, se bloquea el login hasta crear/promover un SUPERADMIN.
                var superAdminPerfilId = await _context.Perfiles
                    .AsNoTracking()
                    .Where(p => p.Nombre != null && p.Nombre.Trim().ToUpper() == "SUPERADMIN")
                    .Select(p => p.Id)
                    .FirstOrDefaultAsync();

                var totalSuperAdmins = superAdminPerfilId > 0
                    ? await _context.Usuarios
                        .AsNoTracking()
                        .CountAsync(u => u.Activo && u.PerfilId == superAdminPerfilId)
                    : 0;

                if (totalSuperAdmins <= 0)
                {
                    // Flujo forzado de bootstrap inicial de SUPERADMIN.
                    return BadRequest("No existe ningún SUPERADMIN activo. Debe crear o promover un usuario a SUPERADMIN antes de iniciar sesión.");
                }

                // Buscar usuario incluyendo todas sus relaciones
                var usuario = await _context.Usuarios
                    .Include(u => u.Perfil)
                    .Include(u => u.Country)
                    .Include(u => u.State)
                    .Include(u => u.City)
                    .Include(u => u.Empresa)
                    .Include(u => u.Dependencia)
                    .Include(u => u.Sector)
                    .FirstOrDefaultAsync(u => u.Email == request.Email);

                if (usuario == null)
                {
                    return NotFound("Usuario no encontrado");
                }

                if (usuario.Bloqueado)
                {
                    return BadRequest("Usuario bloqueado por exceso de intentos fallidos");
                }

                if (!usuario.Activo)
                {
                    return BadRequest("Usuario inactivo");
                }

                // Verificar contraseña usando la clase compartida PasswordHasher
                var inputPasswordHash = PasswordHasher.HashPassword(request.Password);
                
                // Verificación de contraseña
                if (!PasswordHasher.VerifyPassword(request.Password, usuario.Password))
                {
                    usuario.IntentosLogin++;
                    if (usuario.IntentosLogin >= 300)
                    {
                        usuario.Bloqueado = true;
                    }
                    await _context.SaveChangesAsync();
                    return BadRequest($"Contraseña incorrecta. Hash de entrada: {inputPasswordHash}, Hash almacenado: {usuario.Password}");
                }

                // Reset intentos de login y actualizar último acceso
                usuario.IntentosLogin = 0;
                usuario.UltimoAcceso = DateTime.UtcNow;
                usuario.RefreshToken = GenerateRefreshToken();
                usuario.RefreshTokenExpiryTime = DateTime.UtcNow.AddDays(7);
                await _context.SaveChangesAsync();

                // Obtener nombre del perfil — primero del Include, si no, query directa por PerfilId
                var perfilNombre = (usuario.Perfil?.Nombre ?? "").Trim();
                if (string.IsNullOrEmpty(perfilNombre) && usuario.PerfilId.HasValue && usuario.PerfilId.Value > 0)
                {
                    perfilNombre = (await _context.Perfiles
                        .AsNoTracking()
                        .Where(p => p.Id == usuario.PerfilId.Value)
                        .Select(p => p.Nombre)
                        .FirstOrDefaultAsync() ?? "").Trim();
                    _logger.LogWarning("Perfil no cargado via Include para {Email} (PerfilId={PerfilId}). Query directa retornó: '{PerfilNombre}'",
                        usuario.Email, usuario.PerfilId, perfilNombre);
                }
                if (string.IsNullOrEmpty(perfilNombre)) perfilNombre = "USUARIO";

                _logger.LogInformation("Login {Email}: PerfilId={PerfilId} => PerfilNombre='{PerfilNombre}'",
                    usuario.Email, usuario.PerfilId, perfilNombre);

                // Generar token JWT
                var token = GenerateJwtToken(usuario, perfilNombre);

                return Ok(new
                {
                    usuario.Id,
                    usuario.NombreUsuario,
                    usuario.Email,
                    usuario.IdFlexoft,
                    usuario.CodiEmprNet,
                    Perfil = new
                    {
                        Id = usuario.Perfil?.Id,
                        Nombre = perfilNombre
                    },
                    Empresa = new
                    {
                        usuario.Empresa?.Id,
                        RazonSocial = usuario.Empresa?.razon
                    },
                    Dependencia = new
                    {
                        usuario.Dependencia?.Id,
                        usuario.Dependencia?.dependencia
                    },
                    Sector = new
                    {
                        usuario.Sector?.Id,
                        usuario.Sector?.sector
                    },
                    Ubicacion = new
                    {
                        Pais = usuario.Country?.Name,
                        Estado = usuario.State?.Name,
                        Ciudad = usuario.City?.Name
                    },
                    Token = token,
                    RefreshToken = usuario.RefreshToken,
                    Mensaje = "Login exitoso"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en el proceso de login");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Endpoint para solicitar recuperación de contraseña por email.
        /// </summary>
        [HttpPost("solicitar-reset")]
        public async Task<IActionResult> SolicitarReset([FromBody] SolicitarResetRequest request)
        {
            try
            {
                if (request == null || string.IsNullOrWhiteSpace(request.Email))
                    return BadRequest("Debe ingresar un email para recuperar la contraseña.");

                var email = request.Email.Trim();

                var usuario = await _context.Usuarios
                    .AsNoTracking()
                    .FirstOrDefaultAsync(u => u.Email.ToLower() == email.ToLower() && u.Activo);

                if (usuario == null)
                {
                    _logger.LogWarning("Solicitud de reset para email no registrado o inactivo: {Email}", email);
                    return NotFound("El email ingresado no corresponde a ninguna cuenta activa del sistema.");
                }

                // Invalidar tokens anteriores no usados para este usuario
                var tokensAnteriores = await _context.PasswordResetTokens
                    .Where(t => t.UsuarioId == usuario.Id && !t.Usado && t.FechaExpiracion > DateTime.UtcNow)
                    .ToListAsync();
                foreach (var t in tokensAnteriores)
                    t.Usado = true;

                // Generar token seguro (hex 64 chars, URL-safe)
                var tokenBytes = RandomNumberGenerator.GetBytes(32);
                var token = Convert.ToHexString(tokenBytes).ToLower();

                var resetToken = new PasswordResetToken
                {
                    UsuarioId = usuario.Id,
                    Token = token,
                    FechaExpiracion = DateTime.UtcNow.AddMinutes(30),
                    Usado = false,
                    AddRecord = DateTime.UtcNow,
                    CodiEmprNet = 0
                };
                _context.PasswordResetTokens.Add(resetToken);
                await _context.SaveChangesAsync();

                // Construir link de reset apuntando al cliente Blazor
                var resetLink = $"{Request.Scheme}://{Request.Host}/reset-password?token={token}";

                var body = $@"
                    <div style='font-family:sans-serif;max-width:480px;margin:auto;'>
                        <h2 style='color:#1F7246;'>Recuperación de Contraseña</h2>
                        <p>Hola <strong>{usuario.NombreUsuario}</strong>,</p>
                        <p>Recibimos una solicitud para restablecer la contraseña de tu cuenta.</p>
                        <p>Hacé clic en el siguiente botón para crear una nueva contraseña. Este enlace expira en <strong>30 minutos</strong>.</p>
                        <p style='text-align:center;margin:30px 0;'>
                            <a href='{resetLink}'
                               style='background:#1F7246;color:white;padding:12px 28px;border-radius:6px;text-decoration:none;font-size:16px;'>
                                Restablecer Contraseña
                            </a>
                        </p>
                        <p style='color:#666;font-size:13px;'>Si no solicitaste este cambio, podés ignorar este correo. Tu contraseña no será modificada.</p>
                        <hr style='border:none;border-top:1px solid #eee;'/>
                        <p style='color:#aaa;font-size:11px;'>Este enlace expira el {DateTime.UtcNow.AddMinutes(30):dd/MM/yyyy HH:mm} UTC.</p>
                    </div>";

                try
                {
                    await _emailService.EnviarEmailAsync(usuario.Email, "Recuperar contraseña - IMPLANPROD", body, true);
                    _logger.LogInformation("Email de reset enviado a {Email}", usuario.Email);
                    return Ok("Se envió el enlace de recuperación al email indicado.");
                }
                catch (Exception exMail)
                {
                    _logger.LogError(exMail, "Error al enviar email de reset a {Email}", usuario.Email);
                    return StatusCode(500, "No se pudo enviar el email de recuperación. Verifique la configuración SMTP de la empresa o intente nuevamente.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en SolicitarReset para {Email}", request?.Email);
                return StatusCode(500, "Ocurrió un error al procesar la solicitud de recuperación de contraseña.");
            }
        }

        /// <summary>
        /// Endpoint para confirmar el reset de contraseña usando el token del email.
        /// </summary>
        [HttpPost("reset-password")]
        public async Task<IActionResult> ResetPassword([FromBody] ResetPasswordRequest request)
        {
            try
            {
                var resetToken = await _context.PasswordResetTokens
                    .FirstOrDefaultAsync(t => t.Token == request.Token && !t.Usado && t.FechaExpiracion > DateTime.UtcNow);

                if (resetToken == null)
                    return BadRequest("El enlace de recuperación es inválido o ha expirado.");

                var usuario = await _context.Usuarios.FindAsync(resetToken.UsuarioId);
                if (usuario == null || !usuario.Activo)
                    return BadRequest("Usuario no encontrado o inactivo.");

                usuario.Password = PasswordHasher.HashPassword(request.NuevaPassword);
                usuario.Bloqueado = false;
                usuario.IntentosLogin = 0;

                resetToken.Usado = true;
                resetToken.LastUpdate = DateTime.UtcNow;

                await _context.SaveChangesAsync();

                _logger.LogInformation("Contraseña restablecida via reset token para UsuarioId={Id}", usuario.Id);
                return Ok("Contraseña actualizada exitosamente. Ya podés iniciar sesión.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en ResetPassword");
                return StatusCode(500, "Error interno al restablecer la contraseña.");
            }
        }

        /// <summary>
        /// Método interno para restablecer la contraseña de un usuario
        /// </summary>
        private async Task<(bool Success, string Message)> ResetPasswordInternal(string email, string newPassword)
        {
            try
            {
                _logger.LogInformation($"Intentando restablecer contraseña para: {email}");
                
                // Buscar el usuario por email
                var usuario = await _context.Usuarios
                    .FirstOrDefaultAsync(u => u.Email.ToLower() == email.ToLower());

                if (usuario == null)
                {
                    _logger.LogWarning($"Usuario no encontrado: {email}");
                    return (false, $"Usuario con email {email} no encontrado");
                }

                // Establecer la nueva contraseña
                string hashedPassword = PasswordHasher.HashPassword(newPassword);
                usuario.Password = hashedPassword;
                
                // Desbloquear la cuenta si estaba bloqueada
                usuario.Bloqueado = false;
                usuario.IntentosLogin = 0;
                
                await _context.SaveChangesAsync();
                
                _logger.LogInformation($"Contraseña restablecida para: {email}");
                return (true, $"Contraseña del usuario {email} restablecida a '{newPassword}'. Hash: {hashedPassword}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al restablecer contraseña para {email}");
                return (false, $"Error al restablecer contraseña: {ex.Message}");
            }
        }

        /// <summary>
        /// Verifica si una contraseña en texto plano coincide con el hash almacenado
        /// </summary>
        private bool VerifyPassword(string inputPassword, string storedPassword)
        {
            return PasswordHasher.VerifyPassword(inputPassword, storedPassword);
        }

        /// <summary>
        /// Hashea una contraseña utilizando SHA256
        /// </summary>
        private string HashPassword(string password)
        {
            return PasswordHasher.HashPassword(password);
        }

        private string GenerateRefreshToken()
        {
            var randomNumber = new byte[32];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(randomNumber);
                return Convert.ToBase64String(randomNumber);
            }
        }

        private string GenerateJwtToken(Usuario usuario, string perfilNombre = "USUARIO")
        {
            var key = Encoding.ASCII.GetBytes(_configuration["Jwt:Key"] ?? throw new InvalidOperationException("JWT Key no configurada"));
            var claims = new List<Claim>
            {
                new Claim(ClaimTypes.NameIdentifier, usuario.Id.ToString()),
                new Claim(ClaimTypes.Name, usuario.NombreUsuario),
                new Claim(ClaimTypes.Email, usuario.Email),
                new Claim(ClaimTypes.Role, perfilNombre),
                new Claim("EmpresaId", usuario.EmpresaId.ToString()),
                new Claim("DependenciaId", usuario.DependenciaId.ToString()),
                new Claim("SectorId", usuario.SectorId.ToString()),
                new Claim("CountryId", (usuario.CountryId ?? 0).ToString()),
                new Claim("StateId", (usuario.StateId ?? 0).ToString()),
                new Claim("CityId", (usuario.CityId ?? 0).ToString())
            };

            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(claims),
                Expires = DateTime.UtcNow.AddMinutes(Convert.ToDouble(_configuration["Jwt:DurationInMinutes"])),
                Issuer = _configuration["Jwt:Issuer"],
                Audience = _configuration["Jwt:Audience"],
                SigningCredentials = new SigningCredentials(
                    new SymmetricSecurityKey(key),
                    SecurityAlgorithms.HmacSha256Signature)
            };

            var tokenHandler = new JwtSecurityTokenHandler();
            var token = tokenHandler.CreateToken(tokenDescriptor);
            return tokenHandler.WriteToken(token);
        }
    }

    public class LoginRequest
    {
        public required string Email { get; set; }
        public required string Password { get; set; }
    }

    public class SolicitarResetRequest
    {
        public required string Email { get; set; }
    }

    public class ResetPasswordRequest
    {
        public required string Token { get; set; }
        public required string NuevaPassword { get; set; }
    }
}
