using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace PowerPresentations.Core.Rendering;

public sealed class RenderOptions
{
    /// <summary>Ricalcola il rimpicciolimento automatico (normAutofit) invece di usare fontScale memorizzato.</summary>
    public bool RecomputeAutofit { get; init; }
}

/// <summary>
/// Renderer slide: risolve l'ereditarietà slide → layout → master → tema e produce una display list
/// indipendente da WPF (usata da schermo, miniature, presentazione e PDF).
/// </summary>
public sealed partial class SlideRenderer
{
    readonly PptxDocument _doc;
    readonly Dictionary<OpenXmlPart, XElement> _cache = new();
    public SlideRenderer(PptxDocument doc) { _doc = doc; }

    XElement Cached(OpenXmlPart p) { if (!_cache.TryGetValue(p, out var x)) _cache[p] = x = PptxDocument.Xml(p); return x; }

    /// <summary>Stato di rendering di una slide.</summary>
    sealed class St
    {
        public DisplayList Dl = null!; public ColorCtx Cc = null!; public ThemeInfo Theme = null!;
        public XElement Slide = null!; public XElement? Layout, Master, PresDefault, TableStyles;
        public OpenXmlPart Owner = null!; public TextCtx Tc = null!; public RenderOptions Opt = null!;
        public int No;
    }

    readonly record struct Xf(double Sx, double Sy, double Tx, double Ty)
    {
        public static readonly Xf Id = new(1, 1, 0, 0);
        public bool IsId => this == Id;
    }

    static double Emu(double v) => v / 12700.0;

    public DisplayList Render(int index, RenderOptions? opt = null) => Render(_doc.Slides[index], index + 1, opt);

    public DisplayList Render(SlidePart sp, int number, RenderOptions? opt = null)
    {
        var (w, h) = _doc.SlideSize;
        var dl = new DisplayList { Width = w, Height = h };
        var layoutPart = sp.SlideLayoutPart; var masterPart = layoutPart?.SlideMasterPart;
        var themePart = masterPart?.ThemePart;
        var st = new St
        {
            Dl = dl, Slide = PptxDocument.Xml(sp), Layout = layoutPart == null ? null : Cached(layoutPart),
            Master = masterPart == null ? null : Cached(masterPart), Owner = sp, No = number, Opt = opt ?? new RenderOptions()
        };
        st.Theme = ThemeInfo.Parse(themePart == null ? null : Cached(themePart));
        var map = new Dictionary<string, string>(ColorCtx.DefaultMap);
        void Merge(XElement? e) { if (e != null) foreach (var a in e.Attributes()) map[a.Name.LocalName] = a.Value; }
        Merge(st.Master?.Element(Ns.P + "clrMap"));
        Merge(st.Layout?.Element(Ns.P + "clrMapOvr")?.Element(Ns.A + "overrideClrMapping"));
        Merge(st.Slide.Element(Ns.P + "clrMapOvr")?.Element(Ns.A + "overrideClrMapping"));
        st.Cc = new ColorCtx(st.Theme, map);
        var pres = Cached(_doc.PresentationPart);
        st.PresDefault = pres.Element(Ns.P + "defaultTextStyle");
        if (_doc.PresentationPart.TableStylesPart is { } tsp) st.TableStyles = Cached(tsp);
        st.Tc = new TextCtx { Theme = st.Theme, Colors = st.Cc, SlideNo = number, DefaultColor = st.Cc.Resolve(new XElement(Ns.A + "schemeClr", new XAttribute("val", "tx1"))) ?? Rgba.Black };

        // sfondo: slide → layout → master
        foreach (var (root, part) in new[] { ((XElement?)st.Slide, (OpenXmlPart)sp), (st.Layout, layoutPart), (st.Master, masterPart) })
        {
            var bg = root?.Element(Ns.P + "cSld")?.Element(Ns.P + "bg");
            if (bg == null) continue;
            st.Owner = part!;
            var pr = bg.Element(Ns.P + "bgPr");
            if (pr != null) { dl.Background = ResolveFill(st, FillEl(pr), null); if (FillEl(pr)?.Name.LocalName == "blipFill") AddImage(st, FillEl(pr)!, new RectD(0, 0, w, h), 0, false, false, 0); }
            else if (bg.Element(Ns.P + "bgRef") is { } br)
            {
                var clr = st.Cc.Resolve(br.Elements().FirstOrDefault());
                dl.Background = StyleFill(st, (int?)br.Attribute("idx") ?? 0, clr, true);
            }
            break;
        }
        st.Owner = sp;

        bool showLayoutShapes = (string?)st.Slide.Attribute("showMasterSp") != "0";
        bool showMasterShapes = showLayoutShapes && (string?)st.Layout?.Attribute("showMasterSp") != "0";
        if (showMasterShapes && st.Master != null && masterPart != null) { st.Owner = masterPart; DrawTree(st, st.Master.Element(Ns.P + "cSld")?.Element(Ns.P + "spTree"), Xf.Id, false, false); }
        if (showLayoutShapes && st.Layout != null && layoutPart != null) { st.Owner = layoutPart; DrawTree(st, st.Layout.Element(Ns.P + "cSld")?.Element(Ns.P + "spTree"), Xf.Id, false, false); }
        st.Owner = sp;
        DrawTree(st, st.Slide.Element(Ns.P + "cSld")?.Element(Ns.P + "spTree"), Xf.Id, true, true);
        return dl;
    }

    // ---------- helper XML ----------
    static XElement? Ph(XElement sp) => sp.Element(Ns.P + "nvSpPr")?.Element(Ns.P + "nvPr")?.Element(Ns.P + "ph");
    static int IdOf(XElement shape) => (int?)(shape.Elements().FirstOrDefault()?.Element(Ns.P + "cNvPr"))?.Attribute("id") ?? 0;
    static string EffType(XElement ph) => (string?)ph.Attribute("type") ?? "obj";
    static string MasterType(string t) => t == "ctrTitle" ? "title" : t is "subTitle" or "obj" or "tbl" or "chart" or "dgm" or "media" or "clipArt" or "pic" or "body" ? "body" : t;
    static bool SameType(string a, string b) => MasterType(a) == MasterType(b) && (MasterType(a) != "body" || a == b || a is "obj" or "body" || b is "obj" or "body");

    static IEnumerable<XElement> PhShapes(XElement? root) =>
        root?.Element(Ns.P + "cSld")?.Element(Ns.P + "spTree")?.Elements(Ns.P + "sp").Where(s => Ph(s) != null) ?? Enumerable.Empty<XElement>();

    static XElement? FindInLayout(XElement? layout, XElement ph)
    {
        int idx = (int?)ph.Attribute("idx") ?? 0; string t = EffType(ph);
        var all = PhShapes(layout).ToList();
        if (idx != 0) { var m = all.FirstOrDefault(s => ((int?)Ph(s)!.Attribute("idx") ?? 0) == idx); if (m != null) return m; }
        return all.FirstOrDefault(s => SameType(EffType(Ph(s)!), t));
    }

    static XElement? FindInMaster(XElement? master, XElement ph)
    {
        string t = MasterType(EffType(ph));
        return PhShapes(master).FirstOrDefault(s => MasterType(EffType(Ph(s)!)) == t);
    }

    static XElement? FillEl(XElement? pr) => pr?.Elements().FirstOrDefault(e => e.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "blipFill" or "pattFill" or "grpFill");

    // ---------- fill / line ----------
    Fill? ResolveFill(St st, XElement? f, Rgba? ph)
    {
        if (f == null) return null;
        switch (f.Name.LocalName)
        {
            case "solidFill": return st.Cc.Resolve(f.Elements().FirstOrDefault(), ph) is { } c ? new SolidFill(c) : null;
            case "pattFill": return st.Cc.Resolve(f.Element(Ns.A + "fgClr")?.Elements().FirstOrDefault(), ph) is { } p ? new SolidFill(p) : null; // ponytail: motivi → colore di primo piano
            case "gradFill":
                {
                    var stops = new List<GradStop>();
                    foreach (var gs in f.Element(Ns.A + "gsLst")?.Elements(Ns.A + "gs") ?? Enumerable.Empty<XElement>())
                        if (st.Cc.Resolve(gs.Elements().FirstOrDefault(), ph) is { } gc) stops.Add(new GradStop(((double?)gs.Attribute("pos") ?? 0) / 100000.0, gc));
                    if (stops.Count == 0) return null;
                    if (stops.Count == 1) return new SolidFill(stops[0].Color);
                    double ang = ((double?)f.Element(Ns.A + "lin")?.Attribute("ang") ?? 5400000) / 60000.0; // ponytail: gradienti radiali/path → lineari
                    return new GradientFill(stops.OrderBy(s => s.Pos).ToList(), ang);
                }
        }
        return null;
    }

    Fill? StyleFill(St st, int idx, Rgba? ph, bool bg = false)
    {
        if (idx == 0 || st.Theme.Fmt == null) return null;
        var lst = st.Theme.Fmt.Element(Ns.A + (bg ? "bgFillStyleLst" : "fillStyleLst"));
        int i = bg ? idx - 1001 : idx - 1;
        var e = lst?.Elements().ElementAtOrDefault(Math.Max(0, i));
        return ResolveFill(st, e, ph);
    }

    Stroke? ResolveLine(St st, IEnumerable<XElement?> lns, XElement? style)
    {
        var l = lns.Where(x => x != null).Select(x => x!).ToList();
        var fe = l.Select(x => FillEl(x)).FirstOrDefault(x => x != null);
        double? w = l.Select(x => (double?)x.Attribute("w")).FirstOrDefault(x => x != null);
        var dashEl = l.Select(x => x.Element(Ns.A + "prstDash")).FirstOrDefault(x => x != null);
        Rgba? col = null;
        if (fe != null)
        {
            if (fe.Name.LocalName == "noFill") return null;
            col = (ResolveFill(st, fe, null) as SolidFill)?.Color ?? (ResolveFill(st, fe, null) as GradientFill)?.Stops[0].Color;
        }
        else if (style?.Element(Ns.A + "lnRef") is { } lr && (int?)lr.Attribute("idx") is > 0 and var idx)
        {
            var clr = st.Cc.Resolve(lr.Elements().FirstOrDefault());
            var ls = st.Theme.Fmt?.Element(Ns.A + "lnStyleLst")?.Elements(Ns.A + "ln").ElementAtOrDefault(idx - 1);
            col = (ResolveFill(st, FillEl(ls), clr) as SolidFill)?.Color ?? clr;
            w ??= (double?)ls?.Attribute("w");
            dashEl ??= ls?.Element(Ns.A + "prstDash");
        }
        if (col == null) return null;
        double wp = Emu(w ?? 9525); if (wp < 0.25) wp = 0.25;
        double[]? dash = (string?)dashEl?.Attribute("val") switch
        {
            "dash" or "sysDash" => new[] { 4.0, 3.0 }, "dot" or "sysDot" => new[] { 1.0, 2.0 }, "lgDash" => new[] { 8.0, 3.0 },
            "dashDot" or "sysDashDot" => new[] { 4.0, 3.0, 1.0, 3.0 }, _ => null
        };
        return new Stroke(col.Value, wp, dash);
    }

    // ---------- albero delle forme ----------
    void Add(St st, DrawItem it, int id) => st.Dl.Items.Add(it with { ShapeId = id });

    void DrawTree(St st, XElement? tree, Xf xf, bool slideLevel, bool top)
    {
        if (tree == null) return;
        foreach (var e in Expand(tree))
        {
            try { DrawOne(st, e, xf, slideLevel, top); }
            // file non fidato: un valore malformato in UNA forma non deve impedire il disegno del resto della slide
            catch (Exception ex) when (ex is FormatException or OverflowException or InvalidOperationException or ArgumentException or KeyNotFoundException)
            { st.Dl.Warnings.Add($"forma {IdOf(e)} ignorata: {ex.Message}"); }
        }
    }

    void DrawOne(St st, XElement e, Xf xf, bool slideLevel, bool top)
    {
        {
            switch (e.Name.LocalName)
            {
                case "sp": DrawSp(st, e, xf, slideLevel, top); break;
                case "cxnSp": DrawSp(st, e, xf, slideLevel, top); break;
                case "pic": DrawPic(st, e, xf, top); break;
                case "graphicFrame": DrawFrame(st, e, xf, top); break;
                case "grpSp":
                    {
                        var gx = e.Element(Ns.P + "grpSpPr")?.Element(Ns.A + "xfrm");
                        if (gx == null) { DrawTree(st, e, xf, slideLevel, false); break; }
                        double ox = Num(gx, "off", "x"), oy = Num(gx, "off", "y"), cx = Num(gx, "ext", "cx"), cy = Num(gx, "ext", "cy");
                        double chx = Num(gx, "chOff", "x"), chy = Num(gx, "chOff", "y"), chw = Num(gx, "chExt", "cx"), chh = Num(gx, "chExt", "cy");
                        double kx = chw > 0 ? cx / chw : 1, ky = chh > 0 ? cy / chh : 1;
                        // ponytail: la rotazione/ribaltamento del gruppo non è applicata ai figli
                        var nx = new Xf(kx * xf.Sx, ky * xf.Sy, (ox - chx * kx) * xf.Sx + xf.Tx, (oy - chy * ky) * xf.Sy + xf.Ty);
                        DrawTree(st, e, nx, slideLevel, false);
                        break;
                    }
            }
        }
    }

    static IEnumerable<XElement> Expand(XElement tree)
    {
        foreach (var c in tree.Elements())
        {
            if (c.Name == Ns.Mc + "AlternateContent")
            {
                var sel = c.Element(Ns.Mc + "Fallback") ?? c.Element(Ns.Mc + "Choice");
                if (sel != null) foreach (var s in sel.Elements()) yield return s;
            }
            else yield return c;
        }
    }

    static double Num(XElement x, string child, string attr) => (double?)x.Element(Ns.A + child)?.Attribute(attr) ?? 0;

    RectD? BoxOf(XElement? xfrm, Xf xf, out double rot, out bool fh, out bool fv)
    {
        rot = 0; fh = fv = false;
        if (xfrm == null) return null;
        rot = ((double?)xfrm.Attribute("rot") ?? 0) / 60000.0;
        fh = (string?)xfrm.Attribute("flipH") is "1" or "true"; fv = (string?)xfrm.Attribute("flipV") is "1" or "true";
        double x = Num(xfrm, "off", "x"), y = Num(xfrm, "off", "y"), w = Num(xfrm, "ext", "cx"), h = Num(xfrm, "ext", "cy");
        return new RectD(Emu(x * xf.Sx + xf.Tx), Emu(y * xf.Sy + xf.Ty), Emu(w * xf.Sx), Emu(h * xf.Sy));
    }

    void DrawSp(St st, XElement sp, Xf xf, bool slideLevel, bool top)
    {
        var ph = Ph(sp);
        if (ph != null && !slideLevel) return; // i segnaposto di layout/master non compaiono sulla slide
        int id = IdOf(sp);
        XElement? lsp = null, msp = null;
        if (ph != null && slideLevel)
        {
            lsp = FindInLayout(st.Layout, ph);
            msp = FindInMaster(st.Master, lsp != null ? Ph(lsp)! : ph);
        }
        var chain = new List<XElement> { sp }; if (lsp != null) chain.Add(lsp); if (msp != null) chain.Add(msp);
        var prs = chain.Select(s => s.Element(Ns.P + "spPr")).ToList();
        var xfrm = prs.Select(p => p?.Element(Ns.A + "xfrm")).FirstOrDefault(x => x != null);
        if (xfrm == null) return;
        var box = BoxOf(xfrm, xf, out var rot, out var fh, out var fv)!.Value;
        var style = sp.Element(Ns.P + "style");

        // geometria
        var geomEl = prs.Select(p => p?.Element(Ns.A + "prstGeom") ?? p?.Element(Ns.A + "custGeom")).FirstOrDefault(x => x != null);
        bool isLine = sp.Name.LocalName == "cxnSp";
        List<List<Seg>> paths = new();
        string prst = "rect";
        if (geomEl?.Name.LocalName == "custGeom") paths = Geometry.Custom(geomEl, box.W, box.H);
        else
        {
            prst = (string?)geomEl?.Attribute("prst") ?? "rect";
            var av = geomEl?.Element(Ns.A + "avLst")?.Elements(Ns.A + "gd").ToDictionary(g => (string)g.Attribute("name")!, g => double.TryParse(((string?)g.Attribute("fmla") ?? "").Replace("val ", ""), out var d) ? d : (double?)null);
            paths.Add(Geometry.Preset(prst, box.W, box.H, n => av != null && av.TryGetValue(n, out var d) ? d : null));
        }
        bool open = prst is "line" or "straightConnector1" or "bentConnector3" or "bentConnector2" || isLine;

        // riempimento
        Fill? fill = null; bool picFill = false;
        var fe = prs.Select(p => FillEl(p)).FirstOrDefault(x => x != null);
        Rgba? fontClr = null;
        if (fe != null) { fill = ResolveFill(st, fe, null); picFill = fe.Name.LocalName == "blipFill"; }
        else if (style?.Element(Ns.A + "fillRef") is { } fr && (int?)fr.Attribute("idx") is > 0 and var fi)
            fill = StyleFill(st, fi, st.Cc.Resolve(fr.Elements().FirstOrDefault()));
        if (open) fill = null;
        var stroke = ResolveLine(st, prs.Select(p => p?.Element(Ns.A + "ln")), style);

        foreach (var segs in paths)
        {
            var placed = Geometry.Place(segs, box.X, box.Y, box.W, box.H, rot, fh, fv);
            bool closed = segs.Any(s => s.Kind == SegKind.Close);
            if (fill != null && closed && !picFill || stroke != null)
                Add(st, new PathItem(placed, closed ? (picFill ? null : fill) : null, stroke), id);
        }
        if (picFill && fe != null) AddImage(st, fe, box, rot, fh, fv, id);
        if (open && stroke != null)
        {
            var ln = prs.Select(p => p?.Element(Ns.A + "ln")).FirstOrDefault(x => x != null);
            Arrow(st, ln?.Element(Ns.A + "headEnd"), true, box, rot, fh, fv, stroke, id);
            Arrow(st, ln?.Element(Ns.A + "tailEnd"), false, box, rot, fh, fv, stroke, id);
        }

        // testo
        var tb = sp.Element(Ns.P + "txBody");
        if (top && slideLevel) st.Dl.Hits.Add((id, box, true));
        if (tb == null) return;
        if (style?.Element(Ns.A + "fontRef") is { } fo) fontClr = st.Cc.Resolve(fo.Elements().FirstOrDefault());
        DrawText(st, tb, chain, ph, box, rot, fontClr, id);
    }

    void Arrow(St st, XElement? end, bool head, RectD box, double rot, bool fh, bool fv, Stroke stroke, int id)
    {
        var t = (string?)end?.Attribute("type");
        if (t is null or "none") return;
        // punta alla estremità: head = inizio (0,0), tail = fine (w,h) del tracciato locale
        double w = box.W, h = box.H;
        (double x, double y) p = head ? (0, 0) : (w, h), q = head ? (Math.Max(w, 0.001), Math.Max(h, 0.001)) : (0, 0);
        if (w == 0 && h == 0) return;
        double dx = head ? (w == 0 ? 0 : w) : -(w), dy = head ? (h == 0 ? 0 : h) : -(h);
        double len = Math.Sqrt(dx * dx + dy * dy); if (len == 0) return; dx /= len; dy /= len;
        double sz = Math.Max(4, stroke.Width * 3.5), wd = sz * 0.6;
        var segs = new List<Seg>
        {
            new(SegKind.Move, p.x, p.y), new(SegKind.Line, p.x + dx * sz - dy * wd, p.y + dy * sz + dx * wd),
            new(SegKind.Line, p.x + dx * sz + dy * wd, p.y + dy * sz - dx * wd), new(SegKind.Close)
        };
        Add(st, new PathItem(Geometry.Place(segs, box.X, box.Y, w, h, rot, fh, fv), new SolidFill(stroke.Color), null), id);
    }

    void AddImage(St st, XElement blipFill, RectD box, double rot, bool fh, bool fv, int id)
    {
        var rid = (string?)blipFill.Element(Ns.A + "blip")?.Attribute(Ns.R + "embed");
        byte[]? data = null; string ct = "";
        if (rid != null && st.Owner.TryGetPartById(rid, out var part) && part is ImagePart ip)
        {
            ct = ip.ContentType;
            try { using var s = ip.GetStream(); using var ms = new MemoryStream(); s.CopyTo(ms); data = ms.ToArray(); } catch (IOException) { }
        }
        var sr = blipFill.Element(Ns.A + "srcRect");
        double L = ((double?)sr?.Attribute("l") ?? 0) / 100000, T = ((double?)sr?.Attribute("t") ?? 0) / 100000, R_ = ((double?)sr?.Attribute("r") ?? 0) / 100000, B = ((double?)sr?.Attribute("b") ?? 0) / 100000;
        Add(st, new ImageItem(data, ct, box, new RectD(L, T, Math.Max(0.001, 1 - L - R_), Math.Max(0.001, 1 - T - B)), rot, fh, fv), id);
    }

    void DrawPic(St st, XElement pic, Xf xf, bool top)
    {
        var xfrm = pic.Element(Ns.P + "spPr")?.Element(Ns.A + "xfrm");
        var b = BoxOf(xfrm, xf, out var rot, out var fh, out var fv); if (b == null) return;
        int id = IdOf(pic);
        var bf = pic.Element(Ns.P + "blipFill"); if (bf != null) AddImage(st, bf, b.Value, rot, fh, fv, id);
        var stroke = ResolveLine(st, new[] { pic.Element(Ns.P + "spPr")?.Element(Ns.A + "ln") }, pic.Element(Ns.P + "style"));
        if (stroke != null) Add(st, new PathItem(Geometry.Place(Geometry.Preset("rect", b.Value.W, b.Value.H, _ => null), b.Value.X, b.Value.Y, b.Value.W, b.Value.H, rot, fh, fv), null, stroke), id);
        if (top) st.Dl.Hits.Add((id, b.Value, false));
    }

    void DrawFrame(St st, XElement gf, Xf xf, bool top)
    {
        var x = gf.Element(Ns.P + "xfrm");
        var b = BoxOf(x, xf, out _, out _, out _); if (b == null) return;
        int id = IdOf(gf);
        var data = gf.Element(Ns.A + "graphic")?.Element(Ns.A + "graphicData");
        var uri = (string?)data?.Attribute("uri") ?? "";
        if (top) st.Dl.Hits.Add((id, b.Value, uri.EndsWith("/table")));
        if (uri.EndsWith("/table") && data?.Element(Ns.A + "tbl") is { } tbl) { DrawTable(st, tbl, b.Value, id); return; }
        // ponytail: grafici, SmartArt e oggetti OLE non sono disegnati: segnaposto con etichetta (il file resta intatto al salvataggio)
        var r = Geometry.Preset("rect", b.Value.W, b.Value.H, _ => null);
        Add(st, new PathItem(Geometry.Place(r, b.Value.X, b.Value.Y, b.Value.W, b.Value.H, 0, false, false), new SolidFill(new Rgba(240, 240, 240)), new Stroke(new Rgba(170, 170, 170), 0.75, new[] { 4.0, 3.0 })), id);
        string label = uri.EndsWith("/chart") ? "Grafico" : uri.EndsWith("/diagram") ? "SmartArt" : "Oggetto incorporato";
        Add(st, new TextItem(label + " (non visualizzato)", b.Value.X + 8, b.Value.Y + b.Value.H / 2, "Arial", 12, false, false, new Rgba(120, 120, 120)), id);
    }
}
