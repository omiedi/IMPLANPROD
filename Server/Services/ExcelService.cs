using ClosedXML.Excel;

namespace IMPLANPROD.Server.Services
{
    public class ExcelService
    {
        public byte[] GenerarExcel<T>(
            List<T> datos,
            string nombreHoja = "Reporte",
            Dictionary<string, string>? encabezadosPersonalizados = null,
            List<string>? ordenColumnas = null)
        {
            using var workbook = new XLWorkbook();
            var hoja = workbook.Worksheets.Add(nombreHoja);

            // Si se especifica orden de columnas, usar solo esas; si no, filtrar propiedades simples
            List<string> columnas;
            if (ordenColumnas != null && ordenColumnas.Any())
            {
                columnas = ordenColumnas;
            }
            else
            {
                var propiedades = typeof(T).GetProperties()
                    .Where(p => p.PropertyType.IsValueType ||
                               p.PropertyType == typeof(string) ||
                               p.PropertyType == typeof(DateTime) ||
                               p.PropertyType == typeof(DateTime?))
                    .ToArray();
                columnas = propiedades.Select(p => p.Name).ToList();
            }

            var todasPropiedades = typeof(T).GetProperties();

            // Encabezados
            for (int col = 0; col < columnas.Count; col++)
            {
                var nombre = columnas[col];
                var titulo = encabezadosPersonalizados?.GetValueOrDefault(nombre) ?? nombre;

                hoja.Cell(1, col + 1).Value = titulo;
                hoja.Cell(1, col + 1).Style.Font.Bold = true;
            }

            // Filas
            for (int fila = 0; fila < datos.Count; fila++)
            {
                for (int col = 0; col < columnas.Count; col++)
                {
                    var nombrePropiedad = columnas[col];
                    var propiedad = todasPropiedades.FirstOrDefault(p => p.Name == nombrePropiedad);
                    if (propiedad == null) continue;

                    object? valor = null;
                    try
                    {
                        valor = propiedad.GetValue(datos[fila]);
                    }
                    catch (Exception ex)
                    {
                        // Si hay error al obtener el valor (ej: propiedad de navegación no cargada), usar null
                        Console.WriteLine($"[ExcelService] Error al obtener propiedad {nombrePropiedad}: {ex.Message}");
                        valor = null;
                    }

                    var celda = hoja.Cell(fila + 2, col + 1);
                    if (valor is DateTime fecha)
                    {
                        celda.Value = fecha;
                        celda.Style.DateFormat.Format = "dd/MM/yyyy";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                    else if (valor is TimeSpan timeSpan)
                    {
                        // Formatear TimeSpan como string hh:mm:ss
                        int horas = (int)timeSpan.TotalHours;
                        int minutos = timeSpan.Minutes;
                        int segundos = timeSpan.Seconds;
                        celda.Value = $"{horas:D2}:{minutos:D2}:{segundos:D2}";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                    else if (valor is int entero)
                    {
                        celda.Value = entero;
                        celda.Style.NumberFormat.Format = "#,##0"; // ⬅️ entero sin decimales
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (valor is double dbl)
                    {
                        celda.Value = dbl;
                        celda.Style.NumberFormat.Format = "#,##0.00"; // ⬅️ número con decimales
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (valor is decimal dec)
                    {
                        celda.Value = (double)dec;
                        celda.Style.NumberFormat.Format = "#,##0.00";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (valor is float flt)
                    {
                        celda.Value = (double)flt;
                        celda.Style.NumberFormat.Format = "#,##0.00";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (valor != null)
                    {
                        celda.Value = valor.ToString();
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    }
                    else
                    {
                        celda.Value = string.Empty;
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    }

                    // Aplicar estilo según tipo
                    if (valor is DateTime)
                    {
                        celda.Style.DateFormat.Format = "dd/MM/yyyy";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                    else if (valor is int or float or double or decimal)
                    {
                        celda.Style.NumberFormat.Format = "#,##0.00";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else
                    {
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    }
                }
            }

            hoja.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        /// <summary>
        /// Genera un archivo Excel a partir de una lista de diccionarios
        /// </summary>
        public byte[] GenerarExcel(
            List<Dictionary<string, object?>> datos,
            string nombreHoja = "Reporte",
            Dictionary<string, string>? encabezadosPersonalizados = null,
            List<string>? ordenColumnas = null)
        {
            using var workbook = new XLWorkbook();
            var hoja = workbook.Worksheets.Add(nombreHoja);

            if (datos == null || !datos.Any())
            {
                return Array.Empty<byte>();
            }

            // Obtener las columnas del primer registro
            var columnas = ordenColumnas ?? datos.First().Keys.ToList();

            // Encabezados
            for (int col = 0; col < columnas.Count; col++)
            {
                var nombre = columnas[col];
                var titulo = encabezadosPersonalizados?.GetValueOrDefault(nombre) ?? nombre;

                hoja.Cell(1, col + 1).Value = titulo;
                hoja.Cell(1, col + 1).Style.Font.Bold = true;
                hoja.Cell(1, col + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
            }

            // Filas de datos
            for (int fila = 0; fila < datos.Count; fila++)
            {
                var registro = datos[fila];

                for (int col = 0; col < columnas.Count; col++)
                {
                    var nombreColumna = columnas[col];

                    // Buscar el valor en el diccionario (case-insensitive)
                    object? valor = null;

                    // Buscar con el nombre exacto primero
                    if (registro.ContainsKey(nombreColumna))
                    {
                        valor = registro[nombreColumna];
                    }
                    else
                    {
                        // Buscar case-insensitive
                        var claveEncontrada = registro.Keys.FirstOrDefault(k =>
                            string.Equals(k, nombreColumna, StringComparison.OrdinalIgnoreCase));

                        if (claveEncontrada != null)
                        {
                            valor = registro[claveEncontrada];
                        }
                        else if (encabezadosPersonalizados != null)
                        {
                            // Buscar por el nombre personalizado
                            var claveOriginal = encabezadosPersonalizados.FirstOrDefault(x => x.Value == nombreColumna).Key;
                            if (!string.IsNullOrEmpty(claveOriginal))
                            {
                                claveEncontrada = registro.Keys.FirstOrDefault(k =>
                                    string.Equals(k, claveOriginal, StringComparison.OrdinalIgnoreCase));
                                if (claveEncontrada != null)
                                {
                                    valor = registro[claveEncontrada];
                                }
                            }
                        }
                    }

                    var celda = hoja.Cell(fila + 2, col + 1);

                    // Manejar JsonElement si viene de deserialización JSON
                    if (valor is System.Text.Json.JsonElement jsonElement)
                    {
                        valor = jsonElement.ValueKind switch
                        {
                            System.Text.Json.JsonValueKind.Number => jsonElement.TryGetInt32(out var i) ? i : (object?)jsonElement.GetDouble(),
                            System.Text.Json.JsonValueKind.String => jsonElement.GetString(),
                            System.Text.Json.JsonValueKind.True => true,
                            System.Text.Json.JsonValueKind.False => false,
                            System.Text.Json.JsonValueKind.Null => null,
                            _ => jsonElement.ToString()
                        };
                    }

                    if (valor is DateTime fecha)
                    {
                        celda.Value = fecha;
                        celda.Style.DateFormat.Format = "dd/MM/yyyy";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                    else if (valor is TimeSpan timeSpan)
                    {
                        // Formatear TimeSpan como string hh:mm:ss
                        int horas = (int)timeSpan.TotalHours;
                        int minutos = timeSpan.Minutes;
                        int segundos = timeSpan.Seconds;
                        celda.Value = $"{horas:D2}:{minutos:D2}:{segundos:D2}";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                    else if (valor is int entero)
                    {
                        celda.Value = entero;
                        celda.Style.NumberFormat.Format = "#,##0";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (valor is long largo)
                    {
                        celda.Value = largo;
                        celda.Style.NumberFormat.Format = "#,##0";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (valor is double dbl)
                    {
                        celda.Value = dbl;
                        celda.Style.NumberFormat.Format = "#,##0.00";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (valor is decimal dec)
                    {
                        celda.Value = (double)dec;
                        celda.Style.NumberFormat.Format = "#,##0.00";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (valor is float flt)
                    {
                        celda.Value = (double)flt;
                        celda.Style.NumberFormat.Format = "#,##0.00";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (valor is bool boolean)
                    {
                        celda.Value = boolean ? "Sí" : "No";
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                    else if (valor != null)
                    {
                        var valorStr = valor.ToString();
                        // Intentar parsear fechas en formato string
                        if (DateTime.TryParse(valorStr, out var fechaParseada))
                        {
                            celda.Value = fechaParseada;
                            celda.Style.DateFormat.Format = "dd/MM/yyyy";
                            celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }
                        else
                        {
                            celda.Value = valorStr;
                            celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        }
                    }
                    else
                    {
                        celda.Value = string.Empty;
                        celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    }
                }
            }

            // Ajustar ancho de columnas
            hoja.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }
    }
}
