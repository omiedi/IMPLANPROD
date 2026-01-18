using System;
using System.Security.Cryptography;
using System.Text;

namespace IMPLANPROD.Server.Utils
{
    public static class PasswordHasher
    {
        /// <summary>
        /// Hashea una contraseña utilizando SHA256
        /// </summary>
        public static string HashPassword(string password)
        {
            using (var sha256 = SHA256.Create())
            {
                var hashedBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                return Convert.ToBase64String(hashedBytes);
            }
        }

        /// <summary>
        /// Verifica si una contraseña en texto plano coincide con el hash almacenado
        /// </summary>
        public static bool VerifyPassword(string inputPassword, string storedPassword)
        {
            var hash = HashPassword(inputPassword);
            return hash == storedPassword;
        }
    }
}
