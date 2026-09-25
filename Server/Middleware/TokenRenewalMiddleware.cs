using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;

namespace IMPLANPROD.Server.Middleware
{
    /// <summary>
    /// Middleware que implementa renovación de token con actividad (sliding expiration).
    ///
    /// Funcionamiento:
    ///   - El token JWT tiene una duración de 48 horas (configurable en appsettings.json: Jwt:DurationInMinutes).
    ///   - En cada request autenticado, este middleware verifica cuánto tiempo le queda al token.
    ///   - Si le quedan menos de 24 horas (RenewalThresholdHours), genera un token nuevo con las
    ///     mismas claims y lo envía en el header de respuesta "X-Renewed-Token".
    ///   - El cliente (Repository.cs) captura ese header y guarda el nuevo token en sessionStorage.
    ///   - Mientras el usuario tenga actividad al menos una vez cada 48 horas, la sesión nunca expira.
    ///   - Solo expira si el usuario está 48 horas seguidas sin hacer ninguna llamada al backend.
    ///
    /// Esto reemplaza al antiguo InactivityService (desactivado) y evita que el usuario
    /// tenga que re-loguearse mientras tenga actividad esporádica en el sistema.
    /// </summary>
    public class TokenRenewalMiddleware
    {
        private readonly RequestDelegate _next;
        private readonly IConfiguration _configuration;

        /// <summary>
        /// Umbral en horas: si al token le quedan menos de estas horas, se renueva.
        /// Por defecto 24 horas. Con DurationInMinutes=2880 (48hs), el token se renueva
        /// cuando quedan 24 horas o menos, dando una ventana amplia de renovación.
        /// </summary>
        private const int RenewalThresholdHours = 24;

        public TokenRenewalMiddleware(RequestDelegate next, IConfiguration configuration)
        {
            _next = next;
            _configuration = configuration;
        }

        public async Task InvokeAsync(HttpContext context)
        {
            // Capturar el body del response original para poder agregar headers después
            // Continue processing the request first
            await _next(context);

            // Solo renovar si la respuesta fue exitosa (2xx)
            if (context.Response.StatusCode < 200 || context.Response.StatusCode >= 300)
                return;

            // Solo renovar si hay un usuario autenticado
            var authHeader = context.Request.Headers.Authorization.FirstOrDefault();
            if (string.IsNullOrEmpty(authHeader) || !authHeader.StartsWith("Bearer "))
                return;

            var token = authHeader.Substring("Bearer ".Length).Trim();

            try
            {
                // Parsear el token para ver cuánto tiempo le queda
                var tokenHandler = new JwtSecurityTokenHandler();
                if (!tokenHandler.CanReadToken(token))
                    return;

                var jwtToken = tokenHandler.ReadJwtToken(token);
                var expiration = jwtToken.ValidTo;
                var remaining = expiration - DateTime.UtcNow;

                // Si le quedan menos de RenewalThresholdHours, renovar
                if (remaining.TotalHours <= RenewalThresholdHours)
                {
                    var newToken = GenerateNewToken(jwtToken.Claims);
                    if (!string.IsNullOrEmpty(newToken))
                    {
                        context.Response.Headers["X-Renewed-Token"] = newToken;
                    }
                }
            }
            catch
            {
                // Si falla el parseo, no interrumpir el flujo
            }
        }

        /// <summary>
        /// Genera un nuevo token JWT con las mismas claims del token original
        /// pero con una nueva fecha de expiración (DurationInMinutes desde ahora).
        /// </summary>
        private string GenerateNewToken(IEnumerable<Claim> claims)
        {
            try
            {
                var key = Encoding.ASCII.GetBytes(
                    _configuration["Jwt:Key"] ?? throw new InvalidOperationException("JWT Key no configurada"));

                var tokenDescriptor = new SecurityTokenDescriptor
                {
                    Subject = new ClaimsIdentity(claims),
                    Expires = DateTime.UtcNow.AddMinutes(
                        Convert.ToDouble(_configuration["Jwt:DurationInMinutes"])),
                    Issuer = _configuration["Jwt:Issuer"],
                    Audience = _configuration["Jwt:Audience"],
                    SigningCredentials = new SigningCredentials(
                        new SymmetricSecurityKey(key),
                        SecurityAlgorithms.HmacSha256Signature)
                };

                var tokenHandler = new JwtSecurityTokenHandler();
                var newToken = tokenHandler.CreateToken(tokenDescriptor);
                return tokenHandler.WriteToken(newToken);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[TokenRenewalMiddleware] Error al generar nuevo token: {ex.Message}");
                return string.Empty;
            }
        }
    }
}
