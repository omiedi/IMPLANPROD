using IMPLANPROD.Shared.DTOs.Impresiones;
using Microsoft.AspNetCore.Hosting;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;
using System.Text.Json;
using System.Text.RegularExpressions;
using ZXing;
using ZXing.Common;
using ImageSharpImage = SixLabors.ImageSharp.Image;

namespace IMPLANPROD.Server.Services
{
    public class PdfService
    {
        private readonly IWebHostEnvironment _environment;

        public PdfService(IWebHostEnvironment environment)
        {
            _environment = environment;
        }

        public Task<byte[]> GenerateRemitoFromTemplateAsync(
            RemitoPlantillaDetailDto plantilla,
            RemitoPlantillaPreviewResultDto preview,
            CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (plantilla == null)
            {
                throw new ArgumentNullException(nameof(plantilla));
            }

            if (preview == null)
            {
                throw new ArgumentNullException(nameof(preview));
            }

            var document = new Document();
            document.Info.Title = string.IsNullOrWhiteSpace(plantilla.Nombre)
                ? "Remito"
                : $"Remito - {plantilla.Nombre}";
            var normalStyle = document.Styles[StyleNames.Normal];
            if (normalStyle != null)
            {
                normalStyle.Font.Name = "Arial";
            }

            var design = TryParseDesign(plantilla.LienzoJson);
            var hasDesignedLayout = design?.Elements?.Count > 0;

            var section = document.AddSection();
            if (hasDesignedLayout)
            {
                ApplyDesignerPageSetup(section);
            }
            else
            {
                ApplyDefaultPageSetup(section);
            }

            if (design != null)
            {
                ApplyOrientation(section, design.Orientation);
            }

            var tempCodeImagePaths = new List<string>();
            try
            {
                if (hasDesignedLayout)
                {
                    RenderDesign(document, section, design!, preview.Results, tempCodeImagePaths);
                }
                else
                {
                    BuildTitle(section, plantilla.Nombre);
                    BuildHeader(section, preview.Results);
                    BuildDetail(section, preview.Results);
                }

                var renderer = new PdfDocumentRenderer
                {
                    Document = document
                };

                renderer.RenderDocument();

                using var stream = new MemoryStream();
                renderer.PdfDocument.Save(stream, false);
                return Task.FromResult(stream.ToArray());
            }
            finally
            {
                CleanupTempFiles(tempCodeImagePaths);
            }
        }

        public Task<byte[]> GenerateMillimeterPageAsync(
            string? orientation,
            CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var document = new Document();
            document.Info.Title = "Página milimetrada";

            var normalStyle = document.Styles[StyleNames.Normal];
            if (normalStyle != null)
            {
                normalStyle.Font.Name = "Arial";
            }

            var section = document.AddSection();
            ApplyDesignerPageSetup(section);
            ApplyOrientation(section, orientation);

            DrawMillimeterPage(section, orientation);

            var renderer = new PdfDocumentRenderer
            {
                Document = document
            };

            renderer.RenderDocument();

            using var stream = new MemoryStream();
            renderer.PdfDocument.Save(stream, false);
            return Task.FromResult(stream.ToArray());
        }

        private static void ApplyDefaultPageSetup(Section section)
        {
            section.PageSetup.PageFormat = PageFormat.A4;
            section.PageSetup.TopMargin = Unit.FromCentimeter(1.2);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(1.2);
            section.PageSetup.LeftMargin = Unit.FromCentimeter(1.2);
            section.PageSetup.RightMargin = Unit.FromCentimeter(1.2);
        }

        private static void ApplyDesignerPageSetup(Section section)
        {
            section.PageSetup.PageFormat = PageFormat.A4;
            section.PageSetup.TopMargin = Unit.FromMillimeter(0);
            section.PageSetup.BottomMargin = Unit.FromMillimeter(0);
            section.PageSetup.LeftMargin = Unit.FromMillimeter(0);
            section.PageSetup.RightMargin = Unit.FromMillimeter(0);
        }

        private static void DrawMillimeterPage(Section section, string? orientation)
        {
            var isLandscape = string.Equals(orientation, "Landscape", StringComparison.OrdinalIgnoreCase);
            var pageWidthMm = isLandscape ? 297d : 210d;
            var pageHeightMm = isLandscape ? 210d : 297d;

            var rulerBandMm = 8d;
            var gridStartX = rulerBandMm;
            var gridStartY = rulerBandMm;
            var gridWidthMm = Math.Max(10d, pageWidthMm - rulerBandMm);
            var gridHeightMm = Math.Max(10d, pageHeightMm - rulerBandMm);

            var totalVerticalTicks = (int)Math.Floor(gridWidthMm);
            for (var mm = 0; mm <= totalVerticalTicks; mm++)
            {
                var xMm = gridStartX + mm;
                var isMajor = mm % 10 == 0;
                var isHalf = mm % 5 == 0;

                var lineWidth = isMajor ? 0.18 : (isHalf ? 0.14 : 0.1);
                var lineColor = isMajor ? Colors.Gray : Colors.LightGray;
                AddLineFrame(section, xMm, gridStartY, lineWidth, gridHeightMm, lineWidth, lineColor);

                var tickHeight = isMajor ? rulerBandMm : (isHalf ? 5.5 : 3.5);
                AddLineFrame(section, xMm, rulerBandMm - tickHeight, lineWidth, tickHeight, 0.14, Colors.DarkGray);

                if (isMajor)
                {
                    AddLabel(section, mm.ToString(), xMm - 2.2, 0.8, 4.4, 2.8, 5, ParagraphAlignment.Center);
                }
            }

            var totalHorizontalTicks = (int)Math.Floor(gridHeightMm);
            for (var mm = 0; mm <= totalHorizontalTicks; mm++)
            {
                var yMm = gridStartY + mm;
                var isMajor = mm % 10 == 0;
                var isHalf = mm % 5 == 0;

                var lineWidth = isMajor ? 0.18 : (isHalf ? 0.14 : 0.1);
                var lineColor = isMajor ? Colors.Gray : Colors.LightGray;
                AddLineFrame(section, gridStartX, yMm, gridWidthMm, lineWidth, lineWidth, lineColor);

                var tickWidth = isMajor ? rulerBandMm : (isHalf ? 5.5 : 3.5);
                AddLineFrame(section, rulerBandMm - tickWidth, yMm, tickWidth, lineWidth, 0.14, Colors.DarkGray);

                if (isMajor)
                {
                    AddLabel(section, mm.ToString(), 0.6, yMm - 1.5, rulerBandMm - 1.2, 3, 5, ParagraphAlignment.Right);
                }
            }

            AddLineFrame(section, gridStartX, gridStartY, gridWidthMm, 0.2, 0.2, Colors.Black);
            AddLineFrame(section, gridStartX, gridStartY, 0.2, gridHeightMm, 0.2, Colors.Black);
            AddLineFrame(section, gridStartX, gridStartY + gridHeightMm, gridWidthMm, 0.2, 0.2, Colors.Black);
            AddLineFrame(section, gridStartX + gridWidthMm, gridStartY, 0.2, gridHeightMm, 0.2, Colors.Black);

            AddLabel(section, "mm", 1.2, 1.4, 4.8, 2.8, 5, ParagraphAlignment.Left);
        }

        private static void AddLineFrame(
            Section section,
            double xMm,
            double yMm,
            double widthMm,
            double heightMm,
            double strokeMm,
            Color strokeColor)
        {
            var line = section.AddTextFrame();
            line.RelativeHorizontal = RelativeHorizontal.Page;
            line.RelativeVertical = RelativeVertical.Page;
            line.Left = Unit.FromMillimeter(Math.Max(0, xMm));
            line.Top = Unit.FromMillimeter(Math.Max(0, yMm));
            line.Width = Unit.FromMillimeter(Math.Max(0.1, widthMm));
            line.Height = Unit.FromMillimeter(Math.Max(0.1, heightMm));
            line.MarginLeft = 0;
            line.MarginRight = 0;
            line.MarginTop = 0;
            line.MarginBottom = 0;
            line.FillFormat.Visible = false;
            line.LineFormat.Width = Unit.FromMillimeter(Math.Max(0.05, strokeMm));
            line.LineFormat.Color = strokeColor;
            line.LineFormat.DashStyle = DashStyle.Solid;
        }

        private static void AddLabel(
            Section section,
            string text,
            double xMm,
            double yMm,
            double widthMm,
            double heightMm,
            double fontSize,
            ParagraphAlignment alignment)
        {
            var frame = section.AddTextFrame();
            frame.RelativeHorizontal = RelativeHorizontal.Page;
            frame.RelativeVertical = RelativeVertical.Page;
            frame.Left = Unit.FromMillimeter(Math.Max(0, xMm));
            frame.Top = Unit.FromMillimeter(Math.Max(0, yMm));
            frame.Width = Unit.FromMillimeter(Math.Max(1, widthMm));
            frame.Height = Unit.FromMillimeter(Math.Max(1, heightMm));
            frame.MarginLeft = 0;
            frame.MarginRight = 0;
            frame.MarginTop = 0;
            frame.MarginBottom = 0;
            frame.FillFormat.Visible = false;
            frame.LineFormat.Visible = false;

            var p = frame.AddParagraph(text);
            p.Format.Font.Name = "Arial";
            p.Format.Font.Size = fontSize;
            p.Format.Alignment = alignment;
        }

        private static void BuildTitle(Section section, string? nombrePlantilla)
        {
            var title = section.AddParagraph("REMITO");
            title.Format.Font.Name = "Arial";
            title.Format.Font.Size = 14;
            title.Format.Font.Bold = true;
            title.Format.SpaceAfter = Unit.FromCentimeter(0.3);

            if (!string.IsNullOrWhiteSpace(nombrePlantilla))
            {
                var subtitle = section.AddParagraph(nombrePlantilla.Trim());
                subtitle.Format.Font.Name = "Arial";
                subtitle.Format.Font.Size = 9;
                subtitle.Format.Font.Color = Colors.Gray;
                subtitle.Format.SpaceAfter = Unit.FromCentimeter(0.4);
            }
        }

        private static void BuildHeader(Section section, Dictionary<string, List<Dictionary<string, object?>>> results)
        {
            var headerMovisto = TryGetFirstRow(results, "HeaderMovisto");
            var headerCliente = TryGetFirstRow(results, "HeaderCliente");

            if (headerMovisto.Count == 0 && headerCliente.Count == 0)
            {
                return;
            }

            var table = section.AddTable();
            table.Borders.Width = 0.25;
            table.Format.Font.Name = "Arial";
            table.Format.Font.Size = 8;

            table.AddColumn(Unit.FromCentimeter(3.5));
            table.AddColumn(Unit.FromCentimeter(5.6));
            table.AddColumn(Unit.FromCentimeter(3.5));
            table.AddColumn(Unit.FromCentimeter(5.6));

            AddHeaderRows(table, headerMovisto, "Datos del remito");
            AddHeaderRows(table, headerCliente, "Datos del cliente");

            section.AddParagraph().Format.SpaceAfter = Unit.FromCentimeter(0.25);
        }

        private static void AddHeaderRows(Table table, Dictionary<string, object?> data, string sourceName)
        {
            if (data.Count == 0)
            {
                return;
            }

            var sourceRow = table.AddRow();
            sourceRow.Shading.Color = Colors.LightGray;
            sourceRow.Cells[0].MergeRight = 3;
            sourceRow.Cells[0].AddParagraph(sourceName);
            sourceRow.Cells[0].Format.Font.Bold = true;

            var index = 0;
            foreach (var item in data)
            {
                if (index >= 12)
                {
                    break;
                }

                var row = table.AddRow();
                row.Cells[0].AddParagraph(item.Key);
                row.Cells[0].Format.Font.Bold = true;
                row.Cells[1].AddParagraph(FormatValue(item.Value));

                index++;
            }
        }

        private static void BuildDetail(Section section, Dictionary<string, List<Dictionary<string, object?>>> results)
        {
            if (!results.TryGetValue("DetalleMovisto", out var detalleRows) || detalleRows == null || detalleRows.Count == 0)
            {
                section.AddParagraph("Sin detalle para imprimir.").Format.Font.Size = 9;
                return;
            }

            var columns = detalleRows.First().Keys.Take(8).ToList();
            if (columns.Count == 0)
            {
                section.AddParagraph("Sin detalle para imprimir.").Format.Font.Size = 9;
                return;
            }

            var table = section.AddTable();
            table.Format.Font.Name = "Arial";
            table.Format.Font.Size = 8;
            table.Borders.Width = 0.25;

            var width = 18.0 / columns.Count;
            foreach (var _ in columns)
            {
                table.AddColumn(Unit.FromCentimeter(width));
            }

            var header = table.AddRow();
            header.HeadingFormat = true;
            header.Shading.Color = Colors.LightGray;
            header.Format.Font.Bold = true;

            for (var i = 0; i < columns.Count; i++)
            {
                header.Cells[i].AddParagraph(columns[i]);
            }

            foreach (var item in detalleRows)
            {
                var row = table.AddRow();
                for (var i = 0; i < columns.Count; i++)
                {
                    var columnName = columns[i];
                    item.TryGetValue(columnName, out var value);
                    row.Cells[i].AddParagraph(FormatValue(value));
                }
            }
        }

        private static Dictionary<string, object?> TryGetFirstRow(
            Dictionary<string, List<Dictionary<string, object?>>> results,
            string key)
        {
            if (!results.TryGetValue(key, out var rows) || rows == null || rows.Count == 0)
            {
                return new Dictionary<string, object?>();
            }

            return rows[0];
        }

        private static string FormatValue(object? value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            return value switch
            {
                DateTime dt => dt.ToString("dd/MM/yyyy HH:mm"),
                decimal dec => dec.ToString("N2"),
                double dbl => dbl.ToString("N2"),
                float fl => fl.ToString("N2"),
                _ => value.ToString() ?? string.Empty
            };
        }

        private static LienzoDesignDto? TryParseDesign(string? json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }

            try
            {
                return JsonSerializer.Deserialize<LienzoDesignDto>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });
            }
            catch
            {
                return null;
            }
        }

        private static void ApplyOrientation(Section section, string? orientation)
        {
            if (string.Equals(orientation, "Landscape", StringComparison.OrdinalIgnoreCase))
            {
                section.PageSetup.Orientation = Orientation.Landscape;
            }
            else
            {
                section.PageSetup.Orientation = Orientation.Portrait;
            }
        }

        private void RenderDesign(
            Document document,
            Section section,
            LienzoDesignDto design,
            Dictionary<string, List<Dictionary<string, object?>>> results,
            List<string> tempCodeImagePaths)
        {
            var marginLeftMm = Math.Max(0, design.MarginLeftMm);
            var marginTopMm = Math.Max(0, design.MarginTopMm);

            var nonDetailElements = design.Elements
                .Where(x => !string.Equals(x.Type, "detailTable", StringComparison.OrdinalIgnoreCase))
                .ToList();

            var detailContexts = design.Elements
                .Where(x => string.Equals(x.Type, "detailTable", StringComparison.OrdinalIgnoreCase))
                .Select(x => BuildDetailContext(x, results))
                .Where(x => x != null)
                .Cast<DetailRenderContext>()
                .ToList();

            if (detailContexts.Count == 0)
            {
                RenderElements(section, nonDetailElements, results, marginLeftMm, marginTopMm, tempCodeImagePaths);
                AddPageNumber(section, 1, 1, design.Orientation, design.ShowPageNumbers, design.PageNumberAlign, marginLeftMm, marginTopMm);
                return;
            }

            var totalPages = Math.Max(1, detailContexts.Max(x => x.TotalPages));
            for (var pageIndex = 0; pageIndex < totalPages; pageIndex++)
            {
                var targetSection = pageIndex == 0 ? section : document.AddSection();
                if (pageIndex > 0)
                {
                    ApplyDesignerPageSetup(targetSection);
                    ApplyOrientation(targetSection, design.Orientation);
                }

                RenderElements(targetSection, nonDetailElements, results, marginLeftMm, marginTopMm, tempCodeImagePaths);

                foreach (var detailContext in detailContexts)
                {
                    var pageRows = detailContext.Rows
                        .Skip(pageIndex * detailContext.RowsPerPage)
                        .Take(detailContext.RowsPerPage)
                        .ToList();

                    if (pageRows.Count == 0)
                    {
                        continue;
                    }

                    RenderDetailTableRows(targetSection, detailContext.Element, pageRows, marginLeftMm, marginTopMm);
                }

                AddPageNumber(targetSection, pageIndex + 1, totalPages, design.Orientation, design.ShowPageNumbers, design.PageNumberAlign, marginLeftMm, marginTopMm);
            }
        }

        private void RenderElements(
            Section section,
            List<LienzoElementDto> elements,
            Dictionary<string, List<Dictionary<string, object?>>> results,
            double marginLeftMm,
            double marginTopMm,
            List<string> tempCodeImagePaths)
        {
            foreach (var element in elements)
            {
                if (string.Equals(element.Type, "image", StringComparison.OrdinalIgnoreCase))
                {
                    RenderImage(section, element, marginLeftMm, marginTopMm);
                    continue;
                }

                if (IsShapeElement(element))
                {
                    RenderShape(section, element, marginLeftMm, marginTopMm);
                    continue;
                }

                if (IsCodeElement(element))
                {
                    RenderCode(section, element, results, marginLeftMm, marginTopMm, tempCodeImagePaths);
                    continue;
                }

                var text = ResolveElementText(element, results);

                var frame = section.AddTextFrame();
                frame.RelativeHorizontal = RelativeHorizontal.Page;
                frame.RelativeVertical = RelativeVertical.Page;
                frame.Left = Unit.FromMillimeter(Math.Max(0, element.XMm + marginLeftMm));
                frame.Top = Unit.FromMillimeter(Math.Max(0, element.YMm + marginTopMm));
                frame.Width = Unit.FromMillimeter(Math.Max(1, element.WidthMm));
                frame.Height = Unit.FromMillimeter(Math.Max(1, element.HeightMm));
                frame.MarginLeft = 0;
                frame.MarginRight = 0;
                frame.MarginTop = 0;
                frame.MarginBottom = 0;

                var p = frame.AddParagraph(text);
                p.Format.Font.Name = string.IsNullOrWhiteSpace(element.FontName) ? "Arial" : element.FontName;
                p.Format.Font.Size = element.FontSize > 0 ? element.FontSize : 9;
                p.Format.Alignment = MapAlignment(element.Align);
            }
        }

        private static void AddPageNumber(
            Section section,
            int pageNumber,
            int totalPages,
            string? orientation,
            bool showPageNumbers,
            string? pageNumberAlign,
            double marginLeftMm,
            double marginTopMm)
        {
            if (!showPageNumbers)
            {
                return;
            }

            var isLandscape = string.Equals(orientation, "Landscape", StringComparison.OrdinalIgnoreCase);
            var pageWidthMm = isLandscape ? 297d : 210d;
            var pageHeightMm = isLandscape ? 210d : 297d;

            var align = MapAlignment(pageNumberAlign);
            var leftMm = align switch
            {
                ParagraphAlignment.Left => 10d,
                ParagraphAlignment.Center => (pageWidthMm / 2d) - 16d,
                _ => pageWidthMm - 42d
            };

            var frame = section.AddTextFrame();
            frame.RelativeHorizontal = RelativeHorizontal.Page;
            frame.RelativeVertical = RelativeVertical.Page;
            frame.Left = Unit.FromMillimeter(Math.Max(0, leftMm + marginLeftMm));
            frame.Top = Unit.FromMillimeter(Math.Max(0, (pageHeightMm - 10) + marginTopMm));
            frame.Width = Unit.FromMillimeter(32);
            frame.Height = Unit.FromMillimeter(5);
            frame.MarginLeft = 0;
            frame.MarginRight = 0;
            frame.MarginTop = 0;
            frame.MarginBottom = 0;

            var p = frame.AddParagraph($"Página {pageNumber}/{totalPages}");
            p.Format.Font.Name = "Arial";
            p.Format.Font.Size = 8;
            p.Format.Alignment = align;
        }

        private static DetailRenderContext? BuildDetailContext(
            LienzoElementDto element,
            Dictionary<string, List<Dictionary<string, object?>>> results)
        {
            var datasetName = string.IsNullOrWhiteSpace(element.Binding) ? "DetalleMovisto" : element.Binding.Trim();
            if (!TryGetDatasetRows(results, datasetName, out var rows))
            {
                return null;
            }

            var rowHeightMm = element.DetailRowHeightMm > 0 ? element.DetailRowHeightMm : 6;
            var maxRowsByHeight = Math.Max(1, (int)Math.Floor(element.HeightMm / rowHeightMm));
            var reserveForHeader = element.DetailShowHeader ? 1 : 0;
            maxRowsByHeight = Math.Max(1, maxRowsByHeight - reserveForHeader);
            var rowsPerPage = element.DetailMaxRows > 0
                ? Math.Min(element.DetailMaxRows, maxRowsByHeight)
                : maxRowsByHeight;
            rowsPerPage = Math.Max(1, rowsPerPage);

            var totalPages = Math.Max(1, (int)Math.Ceiling(rows.Count / (double)rowsPerPage));
            return new DetailRenderContext
            {
                Element = element,
                Rows = rows,
                RowsPerPage = rowsPerPage,
                TotalPages = totalPages
            };
        }

        private void RenderImage(Section section, LienzoElementDto element, double marginLeftMm, double marginTopMm)
        {
            var imagePath = ResolveImageFilePath(element.ImagePath);
            if (string.IsNullOrWhiteSpace(imagePath))
            {
                return;
            }

            var image = section.AddImage(imagePath);
            image.RelativeHorizontal = RelativeHorizontal.Page;
            image.RelativeVertical = RelativeVertical.Page;
            image.Left = Unit.FromMillimeter(Math.Max(0, element.XMm + marginLeftMm));
            image.Top = Unit.FromMillimeter(Math.Max(0, element.YMm + marginTopMm));
            image.Width = Unit.FromMillimeter(Math.Max(1, element.WidthMm));
            image.Height = Unit.FromMillimeter(Math.Max(1, element.HeightMm));
            image.LockAspectRatio = false;
        }

        private static bool IsShapeElement(LienzoElementDto element)
        {
            return string.Equals(element.Type, "hline", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(element.Type, "vline", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(element.Type, "rectangle", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCodeElement(LienzoElementDto element)
        {
            return string.Equals(element.Type, "qrcode", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(element.Type, "barcode", StringComparison.OrdinalIgnoreCase);
        }

        private static void CleanupTempFiles(List<string> filePaths)
        {
            if (filePaths == null || filePaths.Count == 0)
            {
                return;
            }

            foreach (var path in filePaths)
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    continue;
                }

                try
                {
                    File.Delete(path);
                }
                catch
                {
                }
            }
        }

        private static void RenderCode(
            Section section,
            LienzoElementDto element,
            Dictionary<string, List<Dictionary<string, object?>>> results,
            double marginLeftMm,
            double marginTopMm,
            List<string> tempCodeImagePaths)
        {
            var payload = ResolveElementText(element, results)?.Trim();
            if (string.IsNullOrWhiteSpace(payload))
            {
                return;
            }

            byte[]? pngBytes;
            if (string.Equals(element.Type, "qrcode", StringComparison.OrdinalIgnoreCase))
            {
                pngBytes = GenerateQrPng(payload, element.WidthMm, element.HeightMm);
            }
            else
            {
                pngBytes = GenerateBarcodePng(payload, element.BarcodeSymbology, element.WidthMm, element.HeightMm);
            }

            if (pngBytes == null || pngBytes.Length == 0)
            {
                return;
            }

            var tempFilePath = Path.Combine(Path.GetTempPath(), $"remito-code-{Guid.NewGuid():N}.png");
            File.WriteAllBytes(tempFilePath, pngBytes);
            tempCodeImagePaths.Add(tempFilePath);

            var image = section.AddImage(tempFilePath);
            image.RelativeHorizontal = RelativeHorizontal.Page;
            image.RelativeVertical = RelativeVertical.Page;
            image.Left = Unit.FromMillimeter(Math.Max(0, element.XMm + marginLeftMm));
            image.Top = Unit.FromMillimeter(Math.Max(0, element.YMm + marginTopMm));
            image.Width = Unit.FromMillimeter(Math.Max(1, element.WidthMm));
            image.Height = Unit.FromMillimeter(Math.Max(1, element.HeightMm));
            image.LockAspectRatio = false;
        }

        private static byte[] GenerateQrPng(string payload, double widthMm, double heightMm)
        {
            var targetPx = Math.Max(120, (int)Math.Round(Math.Max(Math.Max(12, widthMm), Math.Max(12, heightMm)) * 8));
            var writer = new BarcodeWriterPixelData
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new EncodingOptions
                {
                    Width = targetPx,
                    Height = targetPx,
                    Margin = 1,
                    PureBarcode = true
                }
            };

            var pixelData = writer.Write(payload);
            using var image = ImageSharpImage.LoadPixelData<Rgba32>(pixelData.Pixels, pixelData.Width, pixelData.Height);
            using var stream = new MemoryStream();
            image.Save(stream, new PngEncoder());
            return stream.ToArray();
        }

        private static byte[]? GenerateBarcodePng(string payload, string? barcodeSymbology, double widthMm, double heightMm)
        {
            var format = ResolveBarcodeFormat(barcodeSymbology, payload, out var normalizedPayload);
            if (string.IsNullOrWhiteSpace(normalizedPayload))
            {
                return null;
            }

            var widthPx = Math.Max(120, (int)Math.Round(Math.Max(10, widthMm) * 8));
            var heightPx = Math.Max(48, (int)Math.Round(Math.Max(8, heightMm) * 8));
            var writer = new BarcodeWriterPixelData
            {
                Format = format,
                Options = new EncodingOptions
                {
                    Width = widthPx,
                    Height = heightPx,
                    Margin = 0,
                    PureBarcode = true
                }
            };

            var pixelData = writer.Write(normalizedPayload);
            using var image = ImageSharpImage.LoadPixelData<Rgba32>(pixelData.Pixels, pixelData.Width, pixelData.Height);
            using var stream = new MemoryStream();
            image.Save(stream, new PngEncoder());
            return stream.ToArray();
        }

        private static BarcodeFormat ResolveBarcodeFormat(string? barcodeSymbology, string payload, out string normalizedPayload)
        {
            var symbology = string.IsNullOrWhiteSpace(barcodeSymbology)
                ? "Code128"
                : barcodeSymbology.Trim();

            if (string.Equals(symbology, "Ean13", StringComparison.OrdinalIgnoreCase))
            {
                normalizedPayload = new string(payload.Where(char.IsDigit).ToArray());
                if (normalizedPayload.Length == 12 || normalizedPayload.Length == 13)
                {
                    return BarcodeFormat.EAN_13;
                }
            }

            normalizedPayload = payload;
            return BarcodeFormat.CODE_128;
        }

        private static void RenderShape(Section section, LienzoElementDto element, double marginLeftMm, double marginTopMm)
        {
            if (string.Equals(element.Type, "hline", StringComparison.OrdinalIgnoreCase))
            {
                var line = section.AddTextFrame();
                line.RelativeHorizontal = RelativeHorizontal.Page;
                line.RelativeVertical = RelativeVertical.Page;
                line.Left = Unit.FromMillimeter(Math.Max(0, element.XMm + marginLeftMm));
                line.Top = Unit.FromMillimeter(Math.Max(0, element.YMm + marginTopMm));
                line.Width = Unit.FromMillimeter(Math.Max(1, element.WidthMm));
                line.Height = Unit.FromMillimeter(Math.Max(0.2, element.StrokeWidth));
                line.MarginLeft = 0;
                line.MarginRight = 0;
                line.MarginTop = 0;
                line.MarginBottom = 0;
                line.FillFormat.Visible = false;
                ApplyLineFormat(line.LineFormat, element);
                return;
            }

            if (string.Equals(element.Type, "vline", StringComparison.OrdinalIgnoreCase))
            {
                var line = section.AddTextFrame();
                line.RelativeHorizontal = RelativeHorizontal.Page;
                line.RelativeVertical = RelativeVertical.Page;
                line.Left = Unit.FromMillimeter(Math.Max(0, element.XMm + marginLeftMm));
                line.Top = Unit.FromMillimeter(Math.Max(0, element.YMm + marginTopMm));
                line.Width = Unit.FromMillimeter(Math.Max(0.2, element.StrokeWidth));
                line.Height = Unit.FromMillimeter(Math.Max(1, element.HeightMm));
                line.MarginLeft = 0;
                line.MarginRight = 0;
                line.MarginTop = 0;
                line.MarginBottom = 0;
                line.FillFormat.Visible = false;
                ApplyLineFormat(line.LineFormat, element);
                return;
            }

            var rectangle = section.AddTextFrame();
            rectangle.RelativeHorizontal = RelativeHorizontal.Page;
            rectangle.RelativeVertical = RelativeVertical.Page;
            rectangle.Left = Unit.FromMillimeter(Math.Max(0, element.XMm + marginLeftMm));
            rectangle.Top = Unit.FromMillimeter(Math.Max(0, element.YMm + marginTopMm));
            rectangle.Width = Unit.FromMillimeter(Math.Max(1, element.WidthMm));
            rectangle.Height = Unit.FromMillimeter(Math.Max(1, element.HeightMm));
            rectangle.MarginLeft = 0;
            rectangle.MarginRight = 0;
            rectangle.MarginTop = 0;
            rectangle.MarginBottom = 0;
            rectangle.FillFormat.Visible = false;
            ApplyLineFormat(rectangle.LineFormat, element);
        }

        private static void ApplyLineFormat(LineFormat lineFormat, LienzoElementDto element)
        {
            lineFormat.Width = Unit.FromMillimeter(Math.Max(0.2, element.StrokeWidth));
            lineFormat.Color = Colors.Black;
            lineFormat.DashStyle = MapDashStyle(element.StrokeStyle);
        }

        private static DashStyle MapDashStyle(string? strokeStyle)
        {
            if (string.Equals(strokeStyle, "Dash", StringComparison.OrdinalIgnoreCase))
            {
                return DashStyle.Dash;
            }

            if (string.Equals(strokeStyle, "Dot", StringComparison.OrdinalIgnoreCase))
            {
                return DashStyle.Dash;
            }

            if (string.Equals(strokeStyle, "DashDot", StringComparison.OrdinalIgnoreCase))
            {
                return DashStyle.DashDot;
            }

            return DashStyle.Solid;
        }

        private string? ResolveImageFilePath(string? configuredPath)
        {
            if (string.IsNullOrWhiteSpace(configuredPath))
            {
                return null;
            }

            if (Path.IsPathRooted(configuredPath) && File.Exists(configuredPath))
            {
                return configuredPath;
            }

            var normalized = configuredPath.Replace('\\', Path.DirectorySeparatorChar).TrimStart('/', '\\');

            var serverWwwRoot = _environment.WebRootPath;
            if (!string.IsNullOrWhiteSpace(serverWwwRoot))
            {
                var serverCandidate = Path.Combine(serverWwwRoot, normalized);
                if (File.Exists(serverCandidate))
                {
                    return serverCandidate;
                }
            }

            var clientWwwRoot = Path.GetFullPath(Path.Combine(_environment.ContentRootPath, "..", "Client", "wwwroot"));
            var clientCandidate = Path.Combine(clientWwwRoot, normalized);
            if (File.Exists(clientCandidate))
            {
                return clientCandidate;
            }

            return null;
        }

        private sealed class DetailRenderContext
        {
            public LienzoElementDto Element { get; set; } = new();
            public List<Dictionary<string, object?>> Rows { get; set; } = new();
            public int RowsPerPage { get; set; }
            public int TotalPages { get; set; }
        }

        private static void RenderDetailTable(
            Section section,
            LienzoElementDto element,
            Dictionary<string, List<Dictionary<string, object?>>> results)
        {
            var datasetName = string.IsNullOrWhiteSpace(element.Binding) ? "DetalleMovisto" : element.Binding.Trim();
            if (!TryGetDatasetRows(results, datasetName, out var rows) || rows.Count == 0)
            {
                return;
            }

            RenderDetailTableRows(section, element, rows, 0, 0);
        }

        private static void RenderDetailTableRows(
            Section section,
            LienzoElementDto element,
            List<Dictionary<string, object?>> rows,
            double marginLeftMm,
            double marginTopMm)
        {
            if (rows.Count == 0)
            {
                return;
            }

            var columns = ParseColumns(element.DetailColumnsCsv);
            if (columns.Count == 0)
            {
                columns = rows[0].Keys.ToList();
            }

            var columnAlignments = ParseColumnAlignments(element.DetailColumnAlignCsv);

            var rowHeightMm = element.DetailRowHeightMm > 0 ? element.DetailRowHeightMm : 6;
            var maxRowsByHeight = Math.Max(1, (int)Math.Floor(element.HeightMm / rowHeightMm));
            var reserveForHeader = element.DetailShowHeader ? 1 : 0;
            maxRowsByHeight = Math.Max(1, maxRowsByHeight - reserveForHeader);
            var maxRows = element.DetailMaxRows > 0
                ? Math.Min(element.DetailMaxRows, maxRowsByHeight)
                : maxRowsByHeight;

            var dataRows = rows.Take(maxRows).ToList();
            if (dataRows.Count == 0)
            {
                return;
            }

            var columnWidthsMm = ResolveColumnWidthsMm(element.WidthMm, columns.Count, element.DetailColumnWidthsCsv);
            var columnOffsetsMm = BuildColumnOffsetsMm(columnWidthsMm);
            var currentY = element.YMm + marginTopMm;
            var showRowLines = element.DetailRowLines;
            var showColumnLines = element.DetailColumnLines;
            var tableTopMm = currentY;
            var tableLeftMm = element.XMm + marginLeftMm;

            if (showRowLines)
            {
                AddHorizontalDivider(section, tableLeftMm, tableTopMm, element.WidthMm);
            }

            if (element.DetailShowHeader)
            {
                for (var i = 0; i < columns.Count; i++)
                {
                    var headerAlignment = ResolveDetailAlignment(element, columns[i], columns[i], columnAlignments);
                    AddCellTextFrame(
                        section,
                        columns[i],
                        tableLeftMm + columnOffsetsMm[i],
                        currentY,
                        columnWidthsMm[i],
                        rowHeightMm,
                        element,
                        bold: true,
                        alignmentOverride: headerAlignment);
                }

                if (showRowLines)
                {
                    AddHorizontalDivider(section, tableLeftMm, currentY + rowHeightMm, element.WidthMm);
                }

                currentY += rowHeightMm;
            }

            for (var rowIndex = 0; rowIndex < dataRows.Count; rowIndex++)
            {
                var row = dataRows[rowIndex];
                for (var i = 0; i < columns.Count; i++)
                {
                    row.TryGetValue(columns[i], out var value);
                    if (value == null)
                    {
                        var found = row.FirstOrDefault(x => string.Equals(x.Key, columns[i], StringComparison.OrdinalIgnoreCase));
                        value = found.Value;
                    }

                    var cellAlignment = ResolveDetailAlignment(element, columns[i], value, columnAlignments);

                    AddCellTextFrame(
                        section,
                        FormatValue(value),
                        tableLeftMm + columnOffsetsMm[i],
                        currentY,
                        columnWidthsMm[i],
                        rowHeightMm,
                        element,
                        bold: false,
                        alignmentOverride: cellAlignment);
                }

                if (showRowLines && rowIndex < dataRows.Count - 1)
                {
                    AddHorizontalDivider(section, tableLeftMm, currentY + rowHeightMm, element.WidthMm);
                }

                currentY += rowHeightMm;
            }

            if (showRowLines)
            {
                AddHorizontalDivider(section, tableLeftMm, currentY, element.WidthMm);
            }

            if (showColumnLines)
            {
                var tableHeightMm = Math.Max(0.2, currentY - tableTopMm);

                AddVerticalDivider(section, tableLeftMm, tableTopMm, tableHeightMm);

                for (var i = 1; i < columnOffsetsMm.Count; i++)
                {
                    var dividerX = tableLeftMm + columnOffsetsMm[i];
                    AddVerticalDivider(section, dividerX, tableTopMm, tableHeightMm);
                }

                AddVerticalDivider(section, tableLeftMm + element.WidthMm, tableTopMm, tableHeightMm);
            }
        }

        private static void AddCellTextFrame(
            Section section,
            string text,
            double xMm,
            double yMm,
            double widthMm,
            double heightMm,
            LienzoElementDto styleSource,
            bool bold,
            ParagraphAlignment? alignmentOverride)
        {
            var frame = section.AddTextFrame();
            frame.RelativeHorizontal = RelativeHorizontal.Page;
            frame.RelativeVertical = RelativeVertical.Page;
            frame.Left = Unit.FromMillimeter(Math.Max(0, xMm));
            frame.Top = Unit.FromMillimeter(Math.Max(0, yMm));
            frame.Width = Unit.FromMillimeter(Math.Max(1, widthMm));
            frame.Height = Unit.FromMillimeter(Math.Max(1, heightMm));
            frame.MarginLeft = 0;
            frame.MarginRight = 0;
            frame.MarginTop = 0;
            frame.MarginBottom = 0;
            frame.FillFormat.Visible = false;
            frame.LineFormat.Width = Unit.FromMillimeter(0);

            var p = frame.AddParagraph(text);
            p.Format.Font.Name = string.IsNullOrWhiteSpace(styleSource.FontName) ? "Arial" : styleSource.FontName;
            p.Format.Font.Size = styleSource.FontSize > 0 ? styleSource.FontSize : 9;
            p.Format.Font.Bold = bold;
            p.Format.Alignment = alignmentOverride ?? MapAlignment(styleSource.Align);
        }

        private static void AddHorizontalDivider(Section section, double xMm, double yMm, double widthMm)
        {
            var line = section.AddTextFrame();
            line.RelativeHorizontal = RelativeHorizontal.Page;
            line.RelativeVertical = RelativeVertical.Page;
            line.Left = Unit.FromMillimeter(Math.Max(0, xMm));
            line.Top = Unit.FromMillimeter(Math.Max(0, yMm));
            line.Width = Unit.FromMillimeter(Math.Max(1, widthMm));
            line.Height = Unit.FromMillimeter(0.2);
            line.MarginLeft = 0;
            line.MarginRight = 0;
            line.MarginTop = 0;
            line.MarginBottom = 0;
            line.FillFormat.Visible = false;
            line.LineFormat.Width = Unit.FromMillimeter(0.2);
            line.LineFormat.Color = Colors.Black;
            line.LineFormat.DashStyle = DashStyle.Solid;
        }

        private static void AddVerticalDivider(Section section, double xMm, double yMm, double heightMm)
        {
            var line = section.AddTextFrame();
            line.RelativeHorizontal = RelativeHorizontal.Page;
            line.RelativeVertical = RelativeVertical.Page;
            line.Left = Unit.FromMillimeter(Math.Max(0, xMm));
            line.Top = Unit.FromMillimeter(Math.Max(0, yMm));
            line.Width = Unit.FromMillimeter(0.2);
            line.Height = Unit.FromMillimeter(Math.Max(1, heightMm));
            line.MarginLeft = 0;
            line.MarginRight = 0;
            line.MarginTop = 0;
            line.MarginBottom = 0;
            line.FillFormat.Visible = false;
            line.LineFormat.Width = Unit.FromMillimeter(0.2);
            line.LineFormat.Color = Colors.Black;
            line.LineFormat.DashStyle = DashStyle.Solid;
        }

        private static List<string> ParseColumns(string? csv)
        {
            if (string.IsNullOrWhiteSpace(csv))
            {
                return new List<string>();
            }

            return csv
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();
        }

        private static Dictionary<string, string> ParseColumnAlignments(string? csv)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(csv))
            {
                return map;
            }

            var pairs = csv.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var pair in pairs)
            {
                var parts = pair.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                if (parts.Length < 2)
                {
                    continue;
                }

                map[parts[0]] = parts[1];
            }

            return map;
        }

        private static List<double> ResolveColumnWidthsMm(double totalWidthMm, int columnCount, string? csv)
        {
            if (columnCount <= 0)
            {
                return new List<double>();
            }

            var safeTotalWidthMm = Math.Max(1, totalWidthMm);
            var weights = ParseColumnWidthWeights(csv, columnCount);
            var totalWeight = weights.Sum();
            if (totalWeight <= 0)
            {
                return Enumerable.Repeat(safeTotalWidthMm / columnCount, columnCount).ToList();
            }

            return weights
                .Select(weight => safeTotalWidthMm * (weight / totalWeight))
                .ToList();
        }

        private static List<double> ParseColumnWidthWeights(string? csv, int columnCount)
        {
            var weights = new List<double>(columnCount);
            var tokens = string.IsNullOrWhiteSpace(csv)
                ? Array.Empty<string>()
                : csv.Split(',', StringSplitOptions.TrimEntries);

            for (var i = 0; i < columnCount; i++)
            {
                if (i < tokens.Length
                    && double.TryParse(tokens[i], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var widthMm)
                    && widthMm > 0)
                {
                    weights.Add(widthMm);
                    continue;
                }

                weights.Add(1d);
            }

            return weights;
        }

        private static List<double> BuildColumnOffsetsMm(IReadOnlyList<double> columnWidthsMm)
        {
            var offsets = new List<double>(columnWidthsMm.Count);
            var accum = 0d;
            for (var i = 0; i < columnWidthsMm.Count; i++)
            {
                offsets.Add(accum);
                accum += columnWidthsMm[i];
            }

            return offsets;
        }

        private static ParagraphAlignment ResolveDetailAlignment(
            LienzoElementDto element,
            string column,
            object? value,
            Dictionary<string, string> columnAlignments)
        {
            if (columnAlignments.TryGetValue(column, out var configuredAlign))
            {
                return MapAlignment(configuredAlign);
            }

            if (element.DetailAutoAlign)
            {
                return IsNumericObject(value) ? ParagraphAlignment.Right : ParagraphAlignment.Left;
            }

            return MapAlignment(element.Align);
        }

        private static bool IsNumericObject(object? value)
        {
            if (value == null)
            {
                return false;
            }

            return value is byte
                or sbyte
                or short
                or ushort
                or int
                or uint
                or long
                or ulong
                or float
                or double
                or decimal
                || decimal.TryParse(value.ToString(), out _)
                || decimal.TryParse(value.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _);
        }

        private static bool TryGetDatasetRows(
            Dictionary<string, List<Dictionary<string, object?>>> results,
            string datasetName,
            out List<Dictionary<string, object?>> rows)
        {
            if (results.TryGetValue(datasetName, out var directRows) && directRows != null)
            {
                rows = directRows;
                return true;
            }

            var found = results.FirstOrDefault(x => string.Equals(x.Key, datasetName, StringComparison.OrdinalIgnoreCase));
            rows = found.Value ?? new List<Dictionary<string, object?>>();
            return rows.Count > 0;
        }

        private static string ResolveElementText(
            LienzoElementDto element,
            Dictionary<string, List<Dictionary<string, object?>>> results)
        {
            if (!string.IsNullOrWhiteSpace(element.Binding))
            {
                var value = TryResolveBinding(results, element.Binding);
                return ApplyFormattingRules(value, element);
            }

            if (string.IsNullOrWhiteSpace(element.Label))
            {
                return string.Empty;
            }

            var replaced = Regex.Replace(
                element.Label,
                "\\{\\{\\s*(?<token>[^{}]+)\\s*\\}\\}",
                match =>
                {
                    var token = match.Groups["token"].Value;
                    var value = TryResolveBinding(results, token);
                    return ApplyFormattingRules(value, element);
                });

            return ApplyFormattingRules(replaced, element);
        }

        private static string? TryResolveBinding(
            Dictionary<string, List<Dictionary<string, object?>>> results,
            string binding)
        {
            var parts = binding.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (parts.Length == 1)
            {
                return TryResolveFieldOnly(results, parts[0]);
            }

            if (parts.Length < 2)
            {
                return null;
            }

            var datasetName = parts[0];
            var fieldName = parts[1];
            if (!results.TryGetValue(datasetName, out var rows) || rows == null || rows.Count == 0)
            {
                var ds = results.FirstOrDefault(x => string.Equals(x.Key, datasetName, StringComparison.OrdinalIgnoreCase));
                rows = ds.Value;
                if (rows == null || rows.Count == 0)
                {
                    return null;
                }
            }

            var row = rows[0];
            if (!row.TryGetValue(fieldName, out var value))
            {
                var found = row.FirstOrDefault(x => string.Equals(x.Key, fieldName, StringComparison.OrdinalIgnoreCase));
                value = found.Value;
            }

            return FormatValue(value);
        }

        private static string? TryResolveFieldOnly(
            Dictionary<string, List<Dictionary<string, object?>>> results,
            string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
            {
                return null;
            }

            var preferredDatasets = new[] { "HeaderCliente", "HeaderMovisto", "DetalleMovisto" };
            foreach (var dataset in preferredDatasets)
            {
                if (!results.TryGetValue(dataset, out var rows) || rows == null || rows.Count == 0)
                {
                    continue;
                }

                var row = rows[0];
                if (TryGetValueCaseInsensitive(row, fieldName, out var value))
                {
                    return FormatValue(value);
                }
            }

            foreach (var item in results)
            {
                var rows = item.Value;
                if (rows == null || rows.Count == 0)
                {
                    continue;
                }

                var row = rows[0];
                if (TryGetValueCaseInsensitive(row, fieldName, out var value))
                {
                    return FormatValue(value);
                }
            }

            return null;
        }

        private static bool TryGetValueCaseInsensitive(Dictionary<string, object?> row, string fieldName, out object? value)
        {
            if (row.TryGetValue(fieldName, out value))
            {
                return true;
            }

            var found = row.FirstOrDefault(x => string.Equals(x.Key, fieldName, StringComparison.OrdinalIgnoreCase));
            value = found.Value;
            return !string.IsNullOrWhiteSpace(found.Key);
        }

        private static string ApplyFormattingRules(string? value, LienzoElementDto element)
        {
            var text = value ?? string.Empty;
            if (element.MaxLength <= 0 || text.Length <= element.MaxLength)
            {
                return text;
            }

            if (element.NumericOverflowAsterisks && IsNumeric(text))
            {
                return new string('*', text.Length);
            }

            return text[..element.MaxLength];
        }

        private static bool IsNumeric(string text)
        {
            return decimal.TryParse(text, out _)
                   || decimal.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _);
        }

        private static ParagraphAlignment MapAlignment(string? align)
        {
            if (string.Equals(align, "Center", StringComparison.OrdinalIgnoreCase))
            {
                return ParagraphAlignment.Center;
            }

            if (string.Equals(align, "Right", StringComparison.OrdinalIgnoreCase))
            {
                return ParagraphAlignment.Right;
            }

            return ParagraphAlignment.Left;
        }

        private sealed class LienzoDesignDto
        {
            public string? Orientation { get; set; }
            public double MarginLeftMm { get; set; }
            public double MarginTopMm { get; set; }
            public bool ShowPageNumbers { get; set; } = true;
            public string? PageNumberAlign { get; set; } = "Right";
            public List<LienzoElementDto> Elements { get; set; } = new();
        }

        private sealed class LienzoElementDto
        {
            public string Type { get; set; } = "text";
            public string Label { get; set; } = string.Empty;
            public string? Binding { get; set; }
            public double XMm { get; set; }
            public double YMm { get; set; }
            public double WidthMm { get; set; } = 50;
            public double HeightMm { get; set; } = 8;
            public string FontName { get; set; } = "Arial";
            public int FontSize { get; set; } = 9;
            public string Align { get; set; } = "Left";
            public int MaxLength { get; set; }
            public bool NumericOverflowAsterisks { get; set; }
            public string? DetailColumnsCsv { get; set; }
            public int DetailRowHeightMm { get; set; } = 6;
            public int DetailMaxRows { get; set; } = 15;
            public bool DetailShowHeader { get; set; } = true;
            public bool DetailAutoAlign { get; set; } = true;
            public bool DetailRowLines { get; set; }
            public bool DetailColumnLines { get; set; }
            public string? DetailColumnAlignCsv { get; set; }
            public string? DetailColumnWidthsCsv { get; set; }
            public string? ImagePath { get; set; }
            public string BarcodeSymbology { get; set; } = "Code128";
            public string StrokeStyle { get; set; } = "Solid";
            public double StrokeWidth { get; set; } = 0.4;
            public string StrokeColor { get; set; } = "#000000";
        }
    }
}
