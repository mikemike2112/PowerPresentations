using System.Globalization;
using System.Xml.Linq;

namespace PowerPresentations.Core.Rendering;

public static class Ns
{
    public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
}

/// <summary>Tema letto dal file: schema colori, font maggiore/minore, formati (fillStyleLst, lnStyleLst, bgFillStyleLst).</summary>
public sealed class ThemeInfo
{
    public Dictionary<string, Rgba> Colors { get; } = new();
    public string MajorFont { get; private set; } = "Calibri Light";
    public string MinorFont { get; private set; } = "Calibri";
    public XElement? Fmt { get; private set; }

    public static ThemeInfo Parse(XElement? theme)
    {
        var t = new ThemeInfo();
        if (theme == null) { t.Defaults(); return t; }
        var te = theme.Element(Ns.A + "themeElements");
        var cs = te?.Element(Ns.A + "clrScheme");
        if (cs != null)
            foreach (var c in cs.Elements())
            {
                var v = c.Elements().FirstOrDefault(); if (v == null) continue;
                var hex = (string?)v.Attribute("val");
                if (v.Name.LocalName == "sysClr") hex = (string?)v.Attribute("lastClr") ?? hex;
                if (hex is { Length: 6 }) try { t.Colors[c.Name.LocalName] = Rgba.FromHex(hex); } catch (FormatException) { }
            }
        var fs = te?.Element(Ns.A + "fontScheme");
        t.MajorFont = (string?)fs?.Element(Ns.A + "majorFont")?.Element(Ns.A + "latin")?.Attribute("typeface") ?? t.MajorFont;
        t.MinorFont = (string?)fs?.Element(Ns.A + "minorFont")?.Element(Ns.A + "latin")?.Attribute("typeface") ?? t.MinorFont;
        t.Fmt = te?.Element(Ns.A + "fmtScheme");
        t.Defaults();
        return t;
    }

    void Defaults()
    {
        void D(string k, string h) { if (!Colors.ContainsKey(k)) Colors[k] = Rgba.FromHex(h); }
        D("dk1", "000000"); D("lt1", "FFFFFF"); D("dk2", "44546A"); D("lt2", "E7E6E6");
        D("accent1", "4472C4"); D("accent2", "ED7D31"); D("accent3", "A5A5A5"); D("accent4", "FFC000"); D("accent5", "5B9BD5"); D("accent6", "70AD47");
        D("hlink", "0563C1"); D("folHlink", "954F72");
    }
}

/// <summary>Risoluzione dei colori DrawingML (srgbClr, schemeClr + clrMap, sysClr, prstClr e trasformazioni).</summary>
public sealed class ColorCtx
{
    public ThemeInfo Theme { get; }
    public IReadOnlyDictionary<string, string> ClrMap { get; }
    public ColorCtx(ThemeInfo theme, IReadOnlyDictionary<string, string> clrMap) { Theme = theme; ClrMap = clrMap; }

    public static readonly Dictionary<string, string> DefaultMap = new()
    {
        ["bg1"] = "lt1", ["tx1"] = "dk1", ["bg2"] = "lt2", ["tx2"] = "dk2", ["accent1"] = "accent1", ["accent2"] = "accent2",
        ["accent3"] = "accent3", ["accent4"] = "accent4", ["accent5"] = "accent5", ["accent6"] = "accent6", ["hlink"] = "hlink", ["folHlink"] = "folHlink"
    };

    static readonly Dictionary<string, string> Named = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000", ["white"] = "FFFFFF", ["red"] = "FF0000", ["green"] = "008000", ["blue"] = "0000FF", ["yellow"] = "FFFF00",
        ["gray"] = "808080", ["grey"] = "808080", ["orange"] = "FFA500", ["purple"] = "800080", ["cyan"] = "00FFFF", ["magenta"] = "FF00FF",
        ["ltGray"] = "D3D3D3", ["dkGray"] = "A9A9A9", ["dkBlue"] = "00008B", ["dkRed"] = "8B0000", ["dkGreen"] = "006400", ["lime"] = "00FF00"
    };

    /// <summary>Risolve l'elemento colore (srgbClr, schemeClr...) con le sue trasformazioni. phClr = colore del segnaposto di stile.</summary>
    public Rgba? Resolve(XElement? clr, Rgba? phClr = null)
    {
        if (clr == null) return null;
        Rgba? c = null;
        switch (clr.Name.LocalName)
        {
            case "srgbClr": { var v = (string?)clr.Attribute("val"); if (v is { Length: 6 }) try { c = Rgba.FromHex(v); } catch (FormatException) { } break; }
            case "sysClr": { var v = (string?)clr.Attribute("lastClr"); if (v is { Length: 6 }) try { c = Rgba.FromHex(v); } catch (FormatException) { } else c = (string?)clr.Attribute("val") == "windowText" ? Rgba.Black : Rgba.White; break; }
            case "prstClr": { var v = (string?)clr.Attribute("val"); if (v != null && Named.TryGetValue(v, out var h)) c = Rgba.FromHex(h); break; }
            case "schemeClr":
                {
                    var n = (string?)clr.Attribute("val") ?? "";
                    if (n == "phClr") c = phClr;
                    else
                    {
                        var m = ClrMap.TryGetValue(n, out var mm) ? mm : n;
                        if (Theme.Colors.TryGetValue(m, out var tc)) c = tc;
                    }
                    break;
                }
        }
        if (c == null) return null;
        var r = c.Value;
        foreach (var t in clr.Elements())
        {
            double v = ((double?)t.Attribute("val") ?? 0) / 100000.0;
            switch (t.Name.LocalName)
            {
                case "alpha": r = r with { A = (byte)Math.Clamp(Math.Round(v * 255), 0, 255) }; break;
                case "alphaMod": r = r with { A = (byte)Math.Clamp(Math.Round(r.A * v), 0, 255) }; break;
                case "lumMod": r = Hsl(r, (h, s, l) => (h, s, l * v)); break;
                case "lumOff": r = Hsl(r, (h, s, l) => (h, s, l + v)); break;
                case "satMod": r = Hsl(r, (h, s, l) => (h, s * v, l)); break;
                case "satOff": r = Hsl(r, (h, s, l) => (h, s + v, l)); break;
                case "hueOff": r = Hsl(r, (h, s, l) => (h + ((double?)t.Attribute("val") ?? 0) / 60000.0, s, l)); break;
                // tint e shade di Office lavorano in RGB lineare (verificato sul rendering di PowerPoint)
                case "tint": r = r with { R = Lin(r.R, c => 1 - (1 - c) * v), G = Lin(r.G, c => 1 - (1 - c) * v), B = Lin(r.B, c => 1 - (1 - c) * v) }; break;
                case "shade": r = r with { R = Lin(r.R, c => c * v), G = Lin(r.G, c => c * v), B = Lin(r.B, c => c * v) }; break;
            }
        }
        return r;
    }

    static byte Lin(byte c, Func<double, double> f)
    {
        double x = c / 255.0; x = x <= 0.04045 ? x / 12.92 : Math.Pow((x + 0.055) / 1.055, 2.4);
        x = Math.Clamp(f(x), 0, 1); x = x <= 0.0031308 ? x * 12.92 : 1.055 * Math.Pow(x, 1 / 2.4) - 0.055;
        return (byte)Math.Round(x * 255);
    }

    static Rgba Hsl(Rgba c, Func<double, double, double, (double h, double s, double l)> f)
    {
        double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
        double mx = Math.Max(r, Math.Max(g, b)), mn = Math.Min(r, Math.Min(g, b)), l = (mx + mn) / 2, h = 0, s = 0;
        if (mx != mn)
        {
            double d = mx - mn;
            s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
            h = mx == r ? (g - b) / d + (g < b ? 6 : 0) : mx == g ? (b - r) / d + 2 : (r - g) / d + 4;
            h *= 60;
        }
        (h, s, l) = f(h, s, l);
        s = Math.Clamp(s, 0, 1); l = Math.Clamp(l, 0, 1); h = ((h % 360) + 360) % 360;
        double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
        double Ch(double t) { t = ((t % 1) + 1) % 1; return t < 1.0 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2.0 / 3 ? p + (q - p) * (2.0 / 3 - t) * 6 : p; }
        return s == 0 ? new Rgba((byte)Math.Round(l * 255), (byte)Math.Round(l * 255), (byte)Math.Round(l * 255), c.A)
            : new Rgba((byte)Math.Round(Ch(h / 360 + 1.0 / 3) * 255), (byte)Math.Round(Ch(h / 360) * 255), (byte)Math.Round(Ch(h / 360 - 1.0 / 3) * 255), c.A);
    }
}
