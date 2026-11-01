using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace PowerPresentations.Core.Rendering;

internal sealed class RProps
{
    public double? Sz; public bool? B, I, U, Strike; public XElement? Fill; public string? Latin, Cap; public double? Baseline; public bool Hlink;
    public RProps Clone() => (RProps)MemberwiseClone();

    static bool Bool(XAttribute? a) => a != null && (a.Value == "1" || a.Value == "true");

    public void Apply(XElement? e)
    {
        if (e == null) return;
        if (e.Attribute("sz") is { } sz && double.TryParse(sz.Value, out var s)) Sz = s / 100.0;
        if (e.Attribute("b") is { } b) B = Bool(b);
        if (e.Attribute("i") is { } i) I = Bool(i);
        if (e.Attribute("u") is { } u) U = u.Value != "none";
        if (e.Attribute("strike") is { } st) Strike = st.Value != "noStrike";
        if (e.Attribute("cap") is { } cap) Cap = cap.Value;
        if (e.Attribute("baseline") is { } bl && double.TryParse(bl.Value, out var bv)) Baseline = bv / 100000.0;
        foreach (var c in e.Elements())
            switch (c.Name.LocalName)
            {
                case "solidFill": Fill = c; break;
                case "latin": Latin = (string?)c.Attribute("typeface"); break;
                case "hlinkClick": Hlink = true; break;
            }
    }
}

internal sealed class PProps
{
    public double? MarL, Indent, LnPct, LnPts, BefPct, BefPts, AftPct, AftPts, BuSzPct;
    public string? Algn, BuKind, BuChar, BuFont, BuAuto; public int? BuStart; public XElement? BuClr;
    public void Apply(XElement? e)
    {
        if (e == null) return;
        if (e.Attribute("marL") is { } m && double.TryParse(m.Value, out var mv)) MarL = mv / 12700.0;
        if (e.Attribute("indent") is { } n && double.TryParse(n.Value, out var nv)) Indent = nv / 12700.0;
        if (e.Attribute("algn") is { } a) Algn = a.Value;
        foreach (var c in e.Elements())
        {
            var v = c.Elements().FirstOrDefault();
            double val = (double?)v?.Attribute("val") ?? 0;
            switch (c.Name.LocalName)
            {
                case "lnSpc": if (v?.Name.LocalName == "spcPct") { LnPct = val / 100000.0; LnPts = null; } else if (v != null) { LnPts = val / 100.0; LnPct = null; } break;
                case "spcBef": if (v?.Name.LocalName == "spcPct") { BefPct = val / 100000.0; BefPts = null; } else if (v != null) { BefPts = val / 100.0; BefPct = null; } break;
                case "spcAft": if (v?.Name.LocalName == "spcPct") { AftPct = val / 100000.0; AftPts = null; } else if (v != null) { AftPts = val / 100.0; AftPct = null; } break;
                case "buNone": BuKind = "none"; break;
                case "buChar": BuKind = "char"; BuChar = (string?)c.Attribute("char"); break;
                case "buAutoNum": BuKind = "auto"; BuAuto = (string?)c.Attribute("type"); BuStart = (int?)c.Attribute("startAt"); break;
                case "buFont": BuFont = (string?)c.Attribute("typeface"); break;
                case "buClr": BuClr = c.Elements().FirstOrDefault(); break;
                case "buSzPct": BuSzPct = ((double?)c.Attribute("val") ?? 100000) / 100000.0; break;
            }
        }
    }
}

internal sealed class TextCtx
{
    public ThemeInfo Theme = null!; public ColorCtx Colors = null!; public int SlideNo = 1;
    public Rgba DefaultColor;
    public string Font(string? latin) => latin switch { null or "" => Theme.MinorFont, "+mn-lt" => Theme.MinorFont, "+mj-lt" => Theme.MajorFont, _ => latin };
}

internal sealed class Laid
{
    public double Height;
    public List<TextItem> Items = new();
    public List<(double X1, double X2, double Y, double W, Rgba C)> Lines = new();
}

internal static class TextEngine
{
    sealed record Tok(RProps P, string Text, double W, double WTrim, bool Break, double Size, string Font, Rgba Color);

    public static bool HasText(XElement? txBody) => txBody != null && txBody.Descendants(Ns.A + "t").Any(t => t.Value.Length > 0);

    static string Roman(int n) { var s = ""; foreach (var (v, r) in new[] { (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i") }) while (n >= v) { s += r; n -= v; } return s; }

    static string AutoNum(string type, int n)
    {
        string body = type.StartsWith("alphaLc") ? Alpha(n).ToLowerInvariant() : type.StartsWith("alphaUc") ? Alpha(n) : type.StartsWith("romanLc") ? Roman(n) : type.StartsWith("romanUc") ? Roman(n).ToUpperInvariant() : n.ToString();
        return type.EndsWith("ParenBoth") ? $"({body})" : type.EndsWith("ParenR") ? body + ")" : type.EndsWith("Plain") ? body : body + ".";
    }
    static string Alpha(int n) { var s = ""; n--; do { s = (char)('A' + n % 26) + s; n = n / 26 - 1; } while (n >= 0); return s; }

    static string BulletGlyph(string? ch, string? font)
    {
        if (string.IsNullOrEmpty(ch)) return "•";
        // ponytail: i glifi Wingdings/Symbol non si mappano su Unicode → punto elenco generico
        if (font != null && (font.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase) || font.StartsWith("Symbol", StringComparison.OrdinalIgnoreCase)))
            return ch == "ü" ? "✓" : ch == "Ø" ? "➢" : ch is "§" or "q" or "n" ? "▪" : "•";
        return ch;
    }

    static readonly Regex Words = new(@"\S+\s*|\s+", RegexOptions.Compiled);

    /// <summary>Impagina il testo in una larghezza data. Le coordinate dei risultati sono relative all'angolo alto-sinistro del blocco di testo.</summary>
    public static Laid Layout(XElement txBody, IReadOnlyList<XElement?> sources, double width, double scale, double lnRed, TextCtx tc, bool wrap = true, int ignoreFillUpTo = 0)
    {
        var res = new Laid(); double y = 0; bool firstPara = true;
        var counters = new Dictionary<int, int>();
        foreach (var p in txBody.Elements(Ns.A + "p"))
        {
            var pPr = p.Element(Ns.A + "pPr");
            int lvl = Math.Clamp((int?)pPr?.Attribute("lvl") ?? 0, 0, 8);
            var pp = new PProps(); var rd = new RProps();
            int si = 0;
            foreach (var s in sources)
            {
                var lv = s?.Element(Ns.A + $"lvl{lvl + 1}pPr");
                pp.Apply(lv); rd.Apply(lv?.Element(Ns.A + "defRPr"));
                if (++si <= ignoreFillUpTo) rd.Fill = null; // il colore di p:style/fontRef prevale sugli stili di default
            }
            pp.Apply(pPr);

            // token
            var toks = new List<Tok>();
            RProps Resolve(XElement? rPr) { var r = rd.Clone(); r.Apply(rPr); return r; }
            Tok Make(RProps r, string text, bool br = false)
            {
                double sz = (r.Sz ?? 18) * scale * (r.Baseline is { } bl && bl != 0 ? 0.65 : 1);
                if (r.Cap == "all") text = text.ToUpperInvariant();
                var face = FontFace.Get(tc.Font(r.Latin), r.B == true, r.I == true);
                var font = face.Family;
                var color = r.Hlink ? tc.Colors.Resolve(new XElement(Ns.A + "schemeClr", new XAttribute("val", "hlink"))) ?? tc.DefaultColor
                    : tc.Colors.Resolve(r.Fill?.Elements().FirstOrDefault()) ?? tc.DefaultColor;
                return new Tok(r, text, face.Measure(text, sz), face.Measure(text.TrimEnd(), sz), br, sz, font, color);
            }
            foreach (var c in p.Elements())
            {
                switch (c.Name.LocalName)
                {
                    case "r":
                    case "fld":
                        {
                            var r = Resolve(c.Element(Ns.A + "rPr"));
                            if (c.Name.LocalName == "r" && c.Element(Ns.A + "rPr")?.Element(Ns.A + "hlinkClick") != null) { r.Hlink = true; r.U = true; }
                            var t = c.Element(Ns.A + "t")?.Value ?? "";
                            if (c.Name.LocalName == "fld" && (string?)c.Attribute("type") == "slidenum") t = tc.SlideNo.ToString();
                            foreach (Match m in Words.Matches(t.Replace("\v", "")))
                                toks.Add(Make(r, m.Value));
                            break;
                        }
                    case "br": toks.Add(Make(Resolve(c.Element(Ns.A + "rPr")), "", true)); break;
                }
            }
            bool empty = !toks.Any(t => !t.Break && t.Text.Length > 0);
            var endR = Resolve(p.Element(Ns.A + "endParaRPr"));
            var dummy = Make(endR, "");

            double marL = pp.MarL ?? 0, indent = pp.Indent ?? 0;
            bool bullet = !empty && pp.BuKind is "char" or "auto" && !empty;
            if (!empty && pp.BuKind is "char" or "auto") bullet = true; else bullet = false;
            string bText = ""; Tok? bTok = null;
            if (bullet)
            {
                var first = toks.First(t => !t.Break);
                if (pp.BuKind == "auto")
                {
                    foreach (var k in counters.Keys.Where(k => k > lvl).ToList()) counters.Remove(k);
                    counters[lvl] = counters.TryGetValue(lvl, out var cn) ? cn + 1 : (pp.BuStart ?? 1);
                    bText = AutoNum(pp.BuAuto ?? "arabicPeriod", counters[lvl]);
                }
                else bText = BulletGlyph(pp.BuChar, pp.BuFont);
                var br = first.P.Clone(); br.Fill = null; br.B = false; br.U = false; br.Strike = false;
                if (pp.BuFont != null && !pp.BuFont.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase) && !pp.BuFont.StartsWith("Symbol", StringComparison.OrdinalIgnoreCase) && !pp.BuFont.StartsWith("+")) br.Latin = pp.BuFont;
                else if (pp.BuKind == "char") br.Latin = "Arial";
                br.Sz = (first.P.Sz ?? 18) * (pp.BuSzPct ?? 1);
                var t = Make(br, bText);
                var col = pp.BuClr != null ? tc.Colors.Resolve(pp.BuClr) ?? first.Color : first.Color;
                bTok = t with { Color = col };
            }
            else
            {
                foreach (var k in counters.Keys.Where(k => k > lvl).ToList()) counters.Remove(k);
                counters.Remove(lvl);
            }

            double firstX = marL + indent;
            double textX1 = bullet ? (marL > firstX + 0.01 ? Math.Max(marL, firstX + bTok!.W) : firstX + bTok!.W + 4) : Math.Max(0, firstX);
            double restX = bullet ? marL : Math.Max(0, marL);
            if (!bullet) textX1 = Math.Max(0, firstX);

            // righe
            var lines = new List<List<Tok>> { new() };
            double cur = 0, avail = Math.Max(10, width - textX1);
            foreach (var t in toks)
            {
                if (t.Break) { lines.Add(new()); cur = 0; avail = Math.Max(10, width - restX); continue; }
                if (wrap && cur + t.WTrim > avail && lines[^1].Count > 0)
                {
                    lines.Add(new()); cur = 0; avail = Math.Max(10, width - restX);
                }
                lines[^1].Add(t); cur += t.W;
            }

            double lnPct = Math.Max(0.1, (pp.LnPct ?? 1.0) - (pp.LnPts == null ? lnRed : 0));
            bool firstLine = true;
            // spazio prima del paragrafo (ignorato sul primo)
            double fsz = (empty ? dummy.Size : toks.FirstOrDefault(t => !t.Break)?.Size) ?? 18 * scale;
            double lhFirst = fsz * FontFace.Get(tc.Font(rd.Latin), false, false).LineHeightEm;
            if (!firstPara) y += pp.BefPts is { } bp ? bp : (pp.BefPct ?? 0) * lhFirst;
            foreach (var ln in lines)
            {
                double maxLh = 0, maxDesc = 0, maxAsc = 0;
                foreach (var t in ln.Count > 0 ? ln : new List<Tok> { dummy })
                {
                    var face = FontFace.Get(t.Font, t.P.B == true, t.P.I == true);
                    maxLh = Math.Max(maxLh, t.Size * face.LineHeightEm);
                    maxDesc = Math.Max(maxDesc, t.Size * face.DescentEm);
                    maxAsc = Math.Max(maxAsc, t.Size * face.AscentEm);
                }
                double lh = pp.LnPts is { } ep ? ep : maxLh * lnPct;
                double baseline = y + lh - maxDesc - (pp.LnPts != null ? lh * 0.2 - maxDesc : 0);
                double lineW = 0; for (int i = 0; i < ln.Count; i++) lineW += i == ln.Count - 1 ? ln[i].WTrim : ln[i].W;
                double startX = firstLine ? textX1 : restX;
                double room = width - startX;
                double x = pp.Algn switch { "ctr" => startX + (room - lineW) / 2, "r" => startX + room - lineW, _ => startX };
                if (pp.Algn is "ctr" or "r" && firstLine && bullet) x = Math.Max(x, startX);
                if (firstLine && bullet && bTok != null)
                    res.Items.Add(new TextItem(bText, firstX, baseline, bTok.Font, bTok.Size, false, false, bTok.Color));
                // unisce i token consecutivi con la stessa run
                for (int i = 0; i < ln.Count;)
                {
                    int j = i; var sb = new System.Text.StringBuilder(); double w = 0;
                    while (j < ln.Count && ReferenceEquals(ln[j].P, ln[i].P)) { sb.Append(ln[j].Text); w += ln[j].W; j++; }
                    var t0 = ln[i]; var txt = sb.ToString(); bool last = j == ln.Count;
                    if (last) { var tr = txt.TrimEnd(); w -= FontFace.Get(t0.Font, t0.P.B == true, t0.P.I == true).Measure(txt[tr.Length..], t0.Size); txt = tr; }
                    double by = baseline - (t0.P.Baseline ?? 0) * (t0.Size / 0.65) ;
                    if (txt.Length > 0)
                    {
                        res.Items.Add(new TextItem(txt, x, by, t0.Font, t0.Size, t0.P.B == true, t0.P.I == true, t0.Color));
                        if (t0.P.U == true) res.Lines.Add((x, x + w, by + t0.Size * 0.1, Math.Max(0.5, t0.Size / 16), t0.Color));
                        if (t0.P.Strike == true) res.Lines.Add((x, x + w, by - t0.Size * 0.3, Math.Max(0.5, t0.Size / 16), t0.Color));
                    }
                    x += w; i = j;
                }
                y += lh; firstLine = false;
            }
            y += pp.AftPts is { } ap ? ap : (pp.AftPct ?? 0) * lhFirst;
            firstPara = false;
        }
        res.Height = y;
        return res;
    }
}
