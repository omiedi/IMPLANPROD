
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
    }

}
