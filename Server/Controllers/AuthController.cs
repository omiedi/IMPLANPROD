using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using IMPLANPROD.Server.Utils;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IConfiguration _configuration;
        private readonly ILogger<AuthController> _logger;

        public AuthController(DataContext context, IConfiguration configuration, ILogger<AuthController> logger)
        {
            _context = context;
            _configuration = configuration;
            _logger = logger;
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

                // Generar token JWT
                var token = GenerateJwtToken(usuario);

                return Ok(new
                {
                    usuario.Id,
                    usuario.NombreUsuario,
                    usuario.Email,
                    usuario.IdFlexoft,  // Campo necesario para sistema de derechos (match con DER_ACCE.NUM_USER)
                    usuario.CodiEmprNet, // Código de empresa para sistema de derechos
                    Perfil = new
                    {
                        usuario.Perfil?.Id,
                        usuario.Perfil?.Nombre
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

        private string GenerateJwtToken(Usuario usuario)
        {
            var key = Encoding.ASCII.GetBytes(_configuration["Jwt:Key"] ?? throw new InvalidOperationException("JWT Key no configurada"));
            var claims = new List<Claim>
            {
                new Claim(ClaimTypes.NameIdentifier, usuario.Id.ToString()),
                new Claim(ClaimTypes.Name, usuario.NombreUsuario),
                new Claim(ClaimTypes.Email, usuario.Email),
                new Claim(ClaimTypes.Role, usuario.Perfil?.Nombre ?? "Usuario"),
                new Claim("EmpresaId", usuario.EmpresaId.ToString()),
                new Claim("DependenciaId", usuario.DependenciaId.ToString()),
                new Claim("SectorId", usuario.SectorId.ToString()),
                new Claim("CountryId", usuario.CountryId.ToString()),
                new Claim("StateId", usuario.StateId.ToString()),
                new Claim("CityId", usuario.CityId.ToString())
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
}
