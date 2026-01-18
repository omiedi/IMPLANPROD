
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace IMPLANPROD.Client.Helpers
{
    public static class ExcelExporter
    {
        static ExcelExporter()
        {
            // Set the license context to non-commercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static byte[] ExportListToExcel<T>(List<T> data)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add header
                var properties = typeof(T).GetProperties();
                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = properties[i].Name;
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }

                // Freeze top row
                worksheet.View.FreezePanes(2, 1);

                // Add data
                for (int i = 0; i < data.Count; i++)
                {
                    for (int j = 0; j < properties.Length; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = properties[j].GetValue(data[i]);
                        worksheet.Cells[i + 2, j + 1].Style.Font.Bold = true;
                    }
                }

                // Auto fit columns
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                return package.GetAsByteArray();
            }
        }
    }
}
