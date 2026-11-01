using System.Xml.Linq;

namespace PowerPresentations.Core.Rendering;

public sealed partial class SlideRenderer
{
    static string? BodyAttr(IEnumerable<XElement?> bps, string name) => bps.Select(b => (string?)b?.Attribute(name)).FirstOrDefault(v => v != null);

    TextCtx CloneTc(St st, Rgba? color) => new() { Theme = st.Theme, Colors = st.Cc, SlideNo = st.No, DefaultColor = color ?? st.Tc.DefaultColor };

    List<XElement?> Sources(St st, List<XElement> chain, XElement? ph)
    {
        var srcs = new List<XElement?> { st.PresDefault };
        if (ph != null)
        {
            var mt = MasterType(EffType(ph));
            srcs.Add(st.Master?.Element(Ns.P + "txStyles")?.Element(Ns.P + (mt == "title" ? "titleStyle" : mt == "body" ? "bodyStyle" : "otherStyle")));
        }
        for (int i = chain.Count - 1; i >= 0; i--) srcs.Add(chain[i].Element(Ns.P + "txBody")?.Element(Ns.A + "lstStyle"));
        return srcs;
    }

    void Emit(St st, Laid laid, double ox, double oy, double rot, RectD box, int id)
    {
        double cx = box.X + box.W / 2, cy = box.Y + box.H / 2;
        foreach (var t in laid.Items)
            Add(st, t with { X = t.X + ox, Baseline = t.Baseline + oy, Rot = rot, Cx = cx, Cy = cy }, id);
        foreach (var l in laid.Lines)
        {
            var segs = new List<Seg> { new(SegKind.Move, l.X1 + ox, l.Y + oy), new(SegKind.Line, l.X2 + ox, l.Y + oy) };
            Add(st, new PathItem(rot == 0 ? segs : RotateAround(segs, rot, cx, cy), null, new Stroke(l.C, l.W)), id);
        }
    }

    static List<Seg> RotateAround(List<Seg> segs, double rot, double cx, double cy)
    {
        double r = rot * Math.PI / 180, c = Math.Cos(r), s = Math.Sin(r);
        (double, double) T(double x, double y) => (cx + (x - cx) * c - (y - cy) * s, cy + (x - cx) * s + (y - cy) * c);
        return segs.Select(g => { var (a, b) = T(g.X1, g.Y1); return g with { X1 = a, Y1 = b }; }).ToList();
    }

    void DrawText(St st, XElement tb, List<XElement> chain, XElement? ph, RectD box, double rot, Rgba? fontClr, int id)
    {
        if (!tb.Elements(Ns.A + "p").Any()) return;
        var bps = chain.Select(s => s.Element(Ns.P + "txBody")?.Element(Ns.A + "bodyPr")).ToList();
        double Ins(string n, double d) => (double.TryParse(BodyAttr(bps, n), out var v) ? v : d) / 12700.0;
        double l = Ins("lIns", 91440), t = Ins("tIns", 45720), r = Ins("rIns", 91440), b = Ins("bIns", 45720);
        double w = Math.Max(1, box.W - l - r), h = Math.Max(1, box.H - t - b);
        bool wrap = BodyAttr(bps, "wrap") != "none";
        var anchor = BodyAttr(bps, "anchor") ?? "t";
        var tc = CloneTc(st, fontClr);
        var src = Sources(st, chain, ph);

        var fit = bps.Select(x => x?.Elements().FirstOrDefault(e => e.Name.LocalName is "normAutofit" or "noAutofit" or "spAutoFit")).FirstOrDefault(x => x != null);
        double scale = 1, red = 0;
        int ign = fontClr != null ? (ph != null ? 2 : 1) : 0;
        var laid = TextEngine.Layout(tb, src, w, 1, 0, tc, wrap, ign);
        if (fit?.Name.LocalName == "normAutofit")
        {
            var fs = (double?)fit.Attribute("fontScale"); var lr = (double?)fit.Attribute("lnSpcReduction");
            if (st.Opt.RecomputeAutofit)
            {
                // passi analoghi a quelli di PowerPoint: 100%, 92,5% (-10% interlinea), poi -7,5% per volta (-20% interlinea)
                var steps = new List<(double, double)> { (1, 0), (0.925, 0.1) }; for (double s = 0.85; s >= 0.25; s -= 0.075) steps.Add((s, 0.2));
                foreach (var (s, rr) in steps)
                {
                    scale = s; red = rr; laid = TextEngine.Layout(tb, src, w, s, rr, tc, wrap, ign);
                    if (laid.Height <= h) break;
                }
                st.Dl.Fits[id] = (scale, red);
            }
            else if (fs != null) { scale = fs.Value / 100000.0; red = (lr ?? 0) / 100000.0; laid = TextEngine.Layout(tb, src, w, scale, red, tc, wrap, ign); }
            // senza fontScale memorizzato PowerPoint non rimpicciolisce finché il testo non viene modificato: stessa regola
        }
        double oy = anchor switch { "ctr" => (h - laid.Height) / 2, "b" => h - laid.Height, _ => 0 };
        Emit(st, laid, box.X + l, box.Y + t + oy, rot, box, id);
    }

    // ---------- tabelle ----------
    const string BuiltInStyleXml = """
<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">
<a:wholeTbl><a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>
<a:left><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:left><a:right><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:right>
<a:top><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top><a:bottom><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom>
<a:insideH><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideH><a:insideV><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideV>
</a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>
<a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H>
<a:band1V><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1V>
<a:lastCol><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastCol>
<a:firstCol><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstCol>
<a:lastRow><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:top><a:ln w="38100"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastRow>
<a:firstRow><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:bottom><a:ln w="38100"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow>
</a:tblStyle>
""";

    static readonly XElement BuiltInStyle = XElement.Parse(BuiltInStyleXml);

    sealed class CellInfo { public XElement Tc = null!; public int R, C, ColSpan = 1, RowSpan = 1; public double W; public Laid? Laid; public double[] M = new double[4]; }

    void DrawTable(St st, XElement tbl, RectD box, int id)
    {
        var cols = tbl.Element(Ns.A + "tblGrid")?.Elements(Ns.A + "gridCol").Select(g => Emu((double?)g.Attribute("w") ?? 0)).ToList() ?? new();
        var trs = tbl.Elements(Ns.A + "tr").ToList();
        if (cols.Count == 0 || trs.Count == 0) return;
        var pr = tbl.Element(Ns.A + "tblPr");
        bool Flag(string n) => (string?)pr?.Attribute(n) is "1" or "true";
        string? sid = pr?.Element(Ns.A + "tableStyleId")?.Value;
        var style = st.TableStyles?.Elements(Ns.A + "tblStyle").FirstOrDefault(s => (string?)s.Attribute("styleId") == sid) ?? BuiltInStyle;
        int nr = trs.Count, nc = cols.Count;
        var cells = new CellInfo?[nr, nc];
        var rowH = trs.Select(t => Emu((double?)t.Attribute("h") ?? 0)).ToArray();
        for (int r = 0; r < nr; r++)
        {
            int c = 0;
            foreach (var tc in trs[r].Elements(Ns.A + "tc"))
            {
                if (c >= nc) break;
                var ci = new CellInfo { Tc = tc, R = r, C = c, ColSpan = Math.Max(1, (int?)tc.Attribute("gridSpan") ?? 1), RowSpan = Math.Max(1, (int?)tc.Attribute("rowSpan") ?? 1) };
                bool merged = (string?)tc.Attribute("hMerge") is "1" or "true" || (string?)tc.Attribute("vMerge") is "1" or "true";
                if (!merged) cells[r, c] = ci;
                c++;
            }
        }
        var tcPrM = new[] { "marL", "marR", "marT", "marB" };
        var defM = new[] { 91440.0, 91440.0, 45720.0, 45720.0 };
        // altezza riga: cresce con il testo (ponytail: celle con rowSpan non allargano le righe)
        for (int r = 0; r < nr; r++)
            for (int c = 0; c < nc; c++)
            {
                var ci = cells[r, c]; if (ci == null) continue;
                ci.W = Enumerable.Range(c, Math.Min(ci.ColSpan, nc - c)).Sum(i => cols[i]);
                var tp = ci.Tc.Element(Ns.A + "tcPr");
                for (int k = 0; k < 4; k++) ci.M[k] = ((double?)tp?.Attribute(tcPrM[k]) ?? defM[k]) / 12700.0;
                var tb = ci.Tc.Element(Ns.A + "txBody");
                if (tb == null) continue;
                var tcx = CloneTc(st, TextColor(st, style, Overlays(style, Flag("firstRow"), Flag("lastRow"), Flag("firstCol"), Flag("lastCol"), Flag("bandRow"), Flag("bandCol"), r, c, nr, nc)));
                ci.Laid = TextEngine.Layout(tb, CellSources(st, style, Flag("firstRow"), Flag("lastRow"), Flag("firstCol"), Flag("lastCol"), Flag("bandRow"), Flag("bandCol"), r, c, nr, nc), Math.Max(1, ci.W - ci.M[0] - ci.M[1]), 1, 0, tcx, true, 1);
                if (ci.RowSpan == 1) rowH[r] = Math.Max(rowH[r], ci.Laid.Height + ci.M[2] + ci.M[3]);
            }
        double y = box.Y;
        var ys = new double[nr + 1]; ys[0] = y; for (int r = 0; r < nr; r++) ys[r + 1] = ys[r] + rowH[r];
        var xs = new double[nc + 1]; xs[0] = box.X; for (int c = 0; c < nc; c++) xs[c + 1] = xs[c] + cols[c];
        // celle e testo
        for (int r = 0; r < nr; r++)
            for (int c = 0; c < nc; c++)
            {
                var ci = cells[r, c]; if (ci == null) continue;
                int c2 = Math.Min(nc, c + ci.ColSpan), r2 = Math.Min(nr, r + ci.RowSpan);
                var cb = new RectD(xs[c], ys[r], xs[c2] - xs[c], ys[r2] - ys[r]);
                var ov = Overlays(style, Flag("firstRow"), Flag("lastRow"), Flag("firstCol"), Flag("lastCol"), Flag("bandRow"), Flag("bandCol"), r, c, nr, nc);
                var tp = ci.Tc.Element(Ns.A + "tcPr");
                Fill? fill = FillEl(tp) is { } fe ? ResolveFill(st, fe, null) : null;
                if (FillEl(tp) == null)
                    foreach (var o in ov)
                    {
                        var ts = o.Element(Ns.A + "tcStyle");
                        var f = FillEl(ts?.Element(Ns.A + "fill"));
                        if (f != null) fill = ResolveFill(st, f, null);
                    }
                var rect = Geometry.Place(Geometry.Preset("rect", cb.W, cb.H, _ => null), cb.X, cb.Y, cb.W, cb.H, 0, false, false);
                if (fill != null) Add(st, new PathItem(rect, fill, null), id);
                if (ci.Laid != null)
                {
                    var tb = ci.Tc.Element(Ns.A + "txBody")!;
                    double avail = cb.H - ci.M[2] - ci.M[3];
                    double oy = ((string?)tp?.Attribute("anchor")) switch { "ctr" => (avail - ci.Laid.Height) / 2, "b" => avail - ci.Laid.Height, _ => 0 };
                    Emit(st, ci.Laid, cb.X + ci.M[0], cb.Y + ci.M[2] + Math.Max(0, oy), 0, cb, id);
                }
                // bordi
                void Edge(string cell, string outer, string inner, bool isOuter, double x1, double y1, double x2, double y2)
                {
                    XElement? ln = tp?.Element(Ns.A + cell);
                    Stroke? sk = null; bool set = false;
                    if (ln != null) { sk = ResolveLine(st, new[] { ln }, null); set = true; }
                    else
                        foreach (var o in ov)
                        {
                            var bd = o.Element(Ns.A + "tcStyle")?.Element(Ns.A + "tcBdr");
                            var e = (o.Name.LocalName == "wholeTbl" ? (isOuter ? bd?.Element(Ns.A + outer) : bd?.Element(Ns.A + inner)) : bd?.Element(Ns.A + outer) ?? (isOuter ? null : bd?.Element(Ns.A + inner)));
                            var lnE = e?.Element(Ns.A + "ln");
                            if (lnE != null) { sk = ResolveLine(st, new[] { lnE }, null); set = true; }
                        }
                    if (set && sk != null)
                        Add(st, new PathItem(new List<Seg> { new(SegKind.Move, x1, y1), new(SegKind.Line, x2, y2) }, null, sk), id);
                }
                Edge("lnL", "left", "insideV", c == 0, cb.X, cb.Y, cb.X, cb.Y + cb.H);
                Edge("lnR", "right", "insideV", c2 == nc, cb.X + cb.W, cb.Y, cb.X + cb.W, cb.Y + cb.H);
                Edge("lnT", "top", "insideH", r == 0, cb.X, cb.Y, cb.X + cb.W, cb.Y);
                Edge("lnB", "bottom", "insideH", r2 == nr, cb.X, cb.Y + cb.H, cb.X + cb.W, cb.Y + cb.H);
            }
    }

    static List<XElement> Overlays(XElement style, bool fr, bool lr, bool fc, bool lc, bool br, bool bc, int r, int c, int nr, int nc)
    {
        var l = new List<XElement>();
        void Add_(string n) { if (style.Element(Ns.A + n) is { } e) l.Add(e); }
        Add_("wholeTbl");
        if (bc) { int s = fc ? 1 : 0; if (c >= s) Add_((c - s) % 2 == 0 ? "band1V" : "band2V"); }
        if (br) { int s = fr ? 1 : 0; if (r >= s) Add_((r - s) % 2 == 0 ? "band1H" : "band2H"); }
        if (fc && c == 0) Add_("firstCol");
        if (lc && c == nc - 1) Add_("lastCol");
        if (fr && r == 0) Add_("firstRow");
        if (lr && r == nr - 1) Add_("lastRow");
        return l;
    }

    Rgba? TextColor(St st, XElement style, List<XElement> ov)
    {
        Rgba? c = null;
        foreach (var o in ov)
        {
            var ts = o.Element(Ns.A + "tcTxStyle"); if (ts == null) continue;
            var ce = ts.Elements().FirstOrDefault(e => e.Name.LocalName is "schemeClr" or "srgbClr" or "sysClr" or "prstClr") ?? ts.Element(Ns.A + "fontRef")?.Elements().FirstOrDefault();
            if (st.Cc.Resolve(ce) is { } rc) c = rc;
        }
        return c;
    }

    List<XElement?> CellSources(St st, XElement style, bool fr, bool lr, bool fc, bool lc, bool br, bool bc, int r, int c, int nr, int nc)
    {
        var ov = Overlays(style, fr, lr, fc, lc, br, bc, r, c, nr, nc);
        bool bold = ov.Any(o => (string?)o.Element(Ns.A + "tcTxStyle")?.Attribute("b") == "on");
        var ls = new XElement(Ns.A + "lstStyle");
        if (bold) for (int i = 1; i <= 9; i++) ls.Add(new XElement(Ns.A + $"lvl{i}pPr", new XElement(Ns.A + "defRPr", new XAttribute("b", "1"))));
        return new List<XElement?> { st.PresDefault, ls };
    }
}
