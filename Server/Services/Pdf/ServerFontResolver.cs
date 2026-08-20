using PdfSharp.Fonts;

namespace IMPLANPROD.Server.Services.Pdf
{
    public sealed class ServerFontResolver : IFontResolver
    {
        private readonly Dictionary<string, byte[]> _fontData;

        public static ServerFontResolver Instance { get; } = new();

        private ServerFontResolver()
        {
            _fontData = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

            TryLoad("Arial#Regular", "arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf");
            TryLoad("Arial#Bold", "arialbd.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf");
            TryLoad("Arial#Italic", "ariali.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf");
            TryLoad("Arial#BoldItalic", "arialbi.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf");

            TryLoad("CourierNew#Regular", "cour.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf");
            TryLoad("CourierNew#Bold", "courbd.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf");
            TryLoad("CourierNew#Italic", "couri.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Oblique.ttf");
            TryLoad("CourierNew#BoldItalic", "courbi.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-BoldOblique.ttf");

            TryLoad("TimesNewRoman#Regular", "times.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf");
            TryLoad("TimesNewRoman#Bold", "timesbd.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf");
            TryLoad("TimesNewRoman#Italic", "timesi.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Italic.ttf");
            TryLoad("TimesNewRoman#BoldItalic", "timesbi.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSerif-BoldItalic.ttf");
        }

        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            var key = ResolveKey(familyName, isBold, isItalic);
            if (_fontData.ContainsKey(key))
            {
                return new FontResolverInfo(key);
            }

            var fallbackKey = ResolveKey("Arial", isBold, isItalic);
            if (_fontData.ContainsKey(fallbackKey))
            {
                return new FontResolverInfo(fallbackKey);
            }

            if (_fontData.ContainsKey("Arial#Regular"))
            {
                return new FontResolverInfo("Arial#Regular");
            }

            return _fontData.Count > 0
                ? new FontResolverInfo(_fontData.Keys.First())
                : null;
        }

        public byte[]? GetFont(string faceName)
        {
            if (_fontData.TryGetValue(faceName, out var data))
            {
                return data;
            }

            if (_fontData.TryGetValue("Arial#Regular", out var fallback))
            {
                return fallback;
            }

            return _fontData.Values.FirstOrDefault();
        }

        private void TryLoad(string faceName, string windowsFileName, string linuxPath)
        {
            var candidates = new List<string>();

            var windowsFontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            if (!string.IsNullOrWhiteSpace(windowsFontsDir))
            {
                candidates.Add(Path.Combine(windowsFontsDir, windowsFileName));
            }

            candidates.Add(linuxPath);

            foreach (var path in candidates)
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    _fontData[faceName] = File.ReadAllBytes(path);
                    return;
                }
            }
        }

        private static string ResolveKey(string familyName, bool isBold, bool isItalic)
        {
            var normalized = (familyName ?? string.Empty).Trim().ToLowerInvariant();
            var familyKey = normalized switch
            {
                "arial" => "Arial",
                "helvetica" => "Arial",
                "sans-serif" => "Arial",
                "courier" => "CourierNew",
                "courier new" => "CourierNew",
                "monospace" => "CourierNew",
                "times" => "TimesNewRoman",
                "times new roman" => "TimesNewRoman",
                "serif" => "TimesNewRoman",
                _ => "Arial"
            };

            var styleKey = isBold && isItalic
                ? "BoldItalic"
                : isBold
                    ? "Bold"
                    : isItalic
                        ? "Italic"
                        : "Regular";

            return $"{familyKey}#{styleKey}";
        }
    }
}
