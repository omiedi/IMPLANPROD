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
        /// Genera el próximo número de documento para un tipo de movimiento específico.
        /// El número se extrae de DoPriCor a partir de la posición 4 en adelante
        /// (formato "TR.NNNNNN", donde "TR." ocupa las 3 primeras posiciones), y no
        /// mediante Split('.'), porque DoPriCor/DoPriLar pueden venir con espacios de
        /// relleno (columnas de ancho fijo heredadas de VB6) que rompían el parseo.
        /// Se calcula el máximo numérico en memoria (y no por orden alfabético de
        /// DoPriCor) porque, al no tener ceros a la izquierda, el orden alfabético
        /// ubicaría "TR.99999" por delante de "TR.100000".
        /// </summary>
        /// <param name="codiMovi">Código de movimiento (ej: 'AJ' para ajustes)</param>
        /// <returns>Número de documento en formato "XX.000001"</returns>
        public async Task<string> GenerateNextDocumentNumberAsync(string codiMovi)
        {
            var codigos = await _context.Movistos
                .Where(m => m.CodiMovi == codiMovi && m.DoPriCor != null)
                .Select(m => m.DoPriCor!)
                .ToListAsync();

            int maxNumero = 0;
            foreach (var codigo in codigos)
            {
                var codigoTrim = codigo.Trim();
                if (codigoTrim.Length < 4)
                {
                    continue;
                }

                // Posición 4 en adelante (índice 3 en base 0) => la parte numérica luego de "XX."
                var parteNumerica = codigoTrim.Substring(3).Trim();
                if (int.TryParse(parteNumerica, out int numero) && numero > maxNumero)
                {
                    maxNumero = numero;
                }
            }

            var numeroSiguiente = maxNumero + 1;
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
