using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;

namespace IMPLANPROD.Server.Utilities
{
    /// <summary>
    /// Clase utilitaria para generar números de documento
    /// </summary>
    public class DocumentNumberGenerator
    {
        private readonly DataContext _context;

        public DocumentNumberGenerator(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Genera el próximo número de documento para un tipo de movimiento específico
        /// </summary>
        /// <param name="codiMovi">Código de movimiento (ej: 'AJ' para ajustes)</param>
        /// <returns>Número de documento en formato "XX.000001"</returns>
        public async Task<string> GenerateNextDocumentNumberAsync(string codiMovi)
        {
            var ultimoMovimiento = await _context.Movistos
                .Where(m => m.CodiMovi == codiMovi)
                .OrderByDescending(m => m.DoPriCor)
                .FirstOrDefaultAsync();

            int numeroSiguiente = 1;
            if (ultimoMovimiento?.DoPriCor != null)
            {
                var ultimoNumero = ultimoMovimiento.DoPriCor.Split('.').Last();
                if (int.TryParse(ultimoNumero, out int numero))
                {
                    numeroSiguiente = numero + 1;
                }
            }

            return $"{codiMovi}.{numeroSiguiente:000000}";
        }

        /// <summary>
        /// Genera el número de documento secundario con la fecha actual
        /// </summary>
        /// <param name="codiMovi">Código de movimiento (ej: 'AJ' para ajustes)</param>
        /// <returns>Número de documento en formato "XX.YYMMDD"</returns>
        public string GenerateSecondaryDocumentNumber(string codiMovi)
        {
            return $"{codiMovi}.{DateTime.Now:yyMMdd}";
        }
    }
}
