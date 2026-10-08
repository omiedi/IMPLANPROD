
namespace IMPLANPROD.Client.Helpers
{
    public static class Formateador
    {
        public static string FormatearHora(DateTime dateTime)
        {
            return dateTime.ToString("HH:mm:ss"); // Formato 24 horas: HH para horas, mm para minutos, ss para segundos.
        }

        public static string FormatearFecha(DateTime dateTime)
        {
            return dateTime.ToString("dd/MM/yyyy"); // Formato de fecha: dd para día, MM para mes, yyyy para año.
        }

        /// <summary>
        /// Símbolo de moneda según Mone_Emis (OCOMENCA/OCOMDETA):
        /// 1 (pesos) → "$", 2 (dólares) → "U$S", otra → descripción TAMONED
        /// si viene informada, sino "$" (pesos, default del sistema).
        /// </summary>
        public static string SimboloMoneda(short? moneEmis, string? descripcionMoneda = null)
        {
            return moneEmis switch
            {
                2 => "U$S",
                > 2 when !string.IsNullOrWhiteSpace(descripcionMoneda) => descripcionMoneda.Trim(),
                _ => "$"
            };
        }

        /// <summary>
        /// Formatea un importe con el símbolo de la moneda de la orden
        /// ("$ 1.234,56" en pesos, "U$S 1.234,56" en dólares).
        /// Reemplaza a ToString("C2"), que toma el símbolo de la cultura del
        /// navegador (podía mostrar € en máquinas configuradas en euros).
        /// </summary>
        public static string FormatearMoneda(decimal? valor, short? moneEmis, string? descripcionMoneda = null)
        {
            return $"{SimboloMoneda(moneEmis, descripcionMoneda)} {(valor ?? 0m).ToString("N2")}";
        }
    }

}
