using Microsoft.Data.SqlClient;
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
        /// Se calcula el máximo numérico en SQL Server con TRY_CAST para evitar transferir
        /// todos los DoPriCor a memoria, lo que provocaba tiempos de respuesta muy altos
        /// cuando hay miles o millones de movimientos.
        /// </summary>
        /// <param name="codiMovi">Código de movimiento (ej: 'AJ' para ajustes)</param>
        /// <returns>Número de documento en formato "XX.000001"</returns>
        public async Task<string> GenerateNextDocumentNumberAsync(string codiMovi)
        {
            var sql = @"
                SELECT MAX(TRY_CAST(SUBSTRING(LTRIM(RTRIM(DoPriCor)), LEN(@codiMovi) + 2, 100) AS INT)) AS [Value]
                FROM movisto
                WHERE CodiMovi = @codiMovi AND DoPriCor IS NOT NULL";

            var parametro = new SqlParameter("@codiMovi", codiMovi);

            var maxNumero = await _context.Database
                .SqlQueryRaw<int?>(sql, parametro)
                .FirstOrDefaultAsync();

            var numeroSiguiente = (maxNumero ?? 0) + 1;
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
