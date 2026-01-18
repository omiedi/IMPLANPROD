using System.Security.Cryptography;
using System.Text;

namespace IMPLANPROD.Server.Utilities
{
    public static class EncryptionHelper
    {
        private static readonly string Key = "M@nt3n1m13nt0ImpK3y2024"; // Clave de 24 caracteres para AES-256

        public static string Encrypt(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            try
            {
                using (Aes aes = Aes.Create())
                {
                    byte[] keyBytes = Encoding.UTF8.GetBytes(Key);
                    Array.Resize(ref keyBytes, 32); // Aseguramos que la clave sea de 32 bytes para AES-256

                    aes.Key = keyBytes;
                    aes.GenerateIV();

                    ICryptoTransform encryptor = aes.CreateEncryptor();

                    using (MemoryStream msEncrypt = new MemoryStream())
                    {
                        // Primero escribimos el IV
                        msEncrypt.Write(aes.IV, 0, aes.IV.Length);

                        using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            swEncrypt.Write(text);
                        }

                        return Convert.ToBase64String(msEncrypt.ToArray());
                    }
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        public static string Decrypt(string encryptedText)
        {
            if (string.IsNullOrEmpty(encryptedText)) return string.Empty;

            try
            {
                byte[] fullCipher = Convert.FromBase64String(encryptedText);

                using (Aes aes = Aes.Create())
                {
                    byte[] keyBytes = Encoding.UTF8.GetBytes(Key);
                    Array.Resize(ref keyBytes, 32);

                    aes.Key = keyBytes;

                    // Extraemos el IV del texto cifrado
                    byte[] iv = new byte[aes.IV.Length];
                    Array.Copy(fullCipher, 0, iv, 0, iv.Length);
                    aes.IV = iv;

                    ICryptoTransform decryptor = aes.CreateDecryptor();

                    using (MemoryStream msDecrypt = new MemoryStream(fullCipher, iv.Length, fullCipher.Length - iv.Length))
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                    {
                        return srDecrypt.ReadToEnd();
                    }
                }
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}

