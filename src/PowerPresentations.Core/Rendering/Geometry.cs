using System.Xml.Linq;

namespace PowerPresentations.Core.Rendering;

/// <summary>Geometrie predefinite (subset) e personalizzate (custGeom) in coordinate locali [0,w]x[0,h].</summary>
public static class Geometry
{
    static Seg M(double x, double y) => new(SegKind.Move, x, y);
    static Seg L(double x, double y) => new(SegKind.Line, x, y);
    static Seg C(double a, double b, double c, double d, double e, double f) => new(SegKind.Cubic, a, b, c, d, e, f);
    static readonly Seg Z = new(SegKind.Close);

    static List<Seg> Poly(params (double x, double y)[] p)
    {
        var l = new List<Seg> { M(p[0].x, p[0].y) };
        for (int i = 1; i < p.Length; i++) l.Add(L(p[i].x, p[i].y));
        l.Add(Z); return l;
    }

    public static List<Seg> Ellipse(double cx, double cy, double rx, double ry)
    {
        const double k = 0.5522847498;
        return new List<Seg>
        {
            M(cx + rx, cy),
            C(cx + rx, cy + k * ry, cx + k * rx, cy + ry, cx, cy + ry),
            C(cx - k * rx, cy + ry, cx - rx, cy + k * ry, cx - rx, cy),
            C(cx - rx, cy - k * ry, cx - k * rx, cy - ry, cx, cy - ry),
            C(cx + k * rx, cy - ry, cx + rx, cy - k * ry, cx + rx, cy), Z
        };
    }

    static List<Seg> RoundRect(double w, double h, double r)
    {
        const double k = 0.5522847498;
        double q = r * (1 - k);
        return new List<Seg>
        {
            M(r, 0), L(w - r, 0), C(w - q, 0, w, q, w, r), L(w, h - r), C(w, h - q, w - q, h, w - r, h),
            L(r, h), C(q, h, 0, h - q, 0, h - r), L(0, r), C(0, q, q, 0, r, 0), Z
        };
    }

    /// <summary>Costruisce il tracciato di un prstGeom con i parametri adj del file.</summary>
    public static List<Seg> Preset(string prst, double w, double h, Func<string, double?> adj)
    {
        double ss = Math.Min(w, h);
        double A(string n, double d) => adj(n) is { } v ? v / 100000.0 : d / 100000.0;
        switch (prst)
        {
            case "ellipse": return Ellipse(w / 2, h / 2, w / 2, h / 2);
            case "roundRect": return RoundRect(w, h, ss * Math.Clamp(A("adj", 16667), 0, 0.5));
            case "triangle": { double a = A("adj", 50000); return Poly((w * a, 0), (w, h), (0, h)); }
            case "rtTriangle": return Poly((0, 0), (w, h), (0, h));
            case "diamond": return Poly((w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2));
            case "parallelogram": { double o = Math.Min(ss * A("adj", 25000), w); return Poly((o, 0), (w, 0), (w - o, h), (0, h)); }
            case "trapezoid": { double o = ss * A("adj", 25000); return Poly((o, 0), (w - o, 0), (w, h), (0, h)); }
            case "pentagon": case "homePlate": { double o = ss * A("adj", 50000); return Poly((0, 0), (w - o, 0), (w, h / 2), (w - o, h), (0, h)); }
            case "chevron": { double o = ss * A("adj", 50000); return Poly((0, 0), (w - o, 0), (w, h / 2), (w - o, h), (0, h), (o, h / 2)); }
            case "hexagon": { double o = ss * A("adj", 25000); return Poly((o, 0), (w - o, 0), (w, h / 2), (w - o, h), (o, h), (0, h / 2)); }
            case "octagon": { double o = ss * A("adj", 29289); return Poly((o, 0), (w - o, 0), (w, o), (w, h - o), (w - o, h), (o, h), (0, h - o), (0, o)); }
            case "plus": case "mathPlus": { double o = ss * A("adj", 25000); return Poly((o, 0), (w - o, 0), (w - o, o), (w, o), (w, h - o), (w - o, h - o), (w - o, h), (o, h), (o, h - o), (0, h - o), (0, o), (o, o)); }
            case "rightArrow": { double t = h * (1 - A("adj1", 50000)) / 2, hd = Math.Min(ss * A("adj2", 50000), w); return Poly((0, t), (w - hd, t), (w - hd, 0), (w, h / 2), (w - hd, h), (w - hd, h - t), (0, h - t)); }
            case "leftArrow": { double t = h * (1 - A("adj1", 50000)) / 2, hd = Math.Min(ss * A("adj2", 50000), w); return Poly((w, t), (hd, t), (hd, 0), (0, h / 2), (hd, h), (hd, h - t), (w, h - t)); }
            case "upArrow": { double t = w * (1 - A("adj1", 50000)) / 2, hd = Math.Min(ss * A("adj2", 50000), h); return Poly((t, h), (t, hd), (0, hd), (w / 2, 0), (w, hd), (w - t, hd), (w - t, h)); }
            case "downArrow": { double t = w * (1 - A("adj1", 50000)) / 2, hd = Math.Min(ss * A("adj2", 50000), h); return Poly((t, 0), (t, h - hd), (0, h - hd), (w / 2, h), (w, h - hd), (w - t, h - hd), (w - t, 0)); }
            case "star5":
                {
                    var p = new (double, double)[10];
                    double cy = h / 2 + h * 0.025, ro = h / 1.9, ri = ro * 0.382, sx = w / h;
                    for (int i = 0; i < 10; i++) { double a = -Math.PI / 2 + i * Math.PI / 5, r = i % 2 == 0 ? ro : ri; p[i] = (w / 2 + Math.Cos(a) * r * sx, cy + Math.Sin(a) * r); }
                    return Poly(p);
                }
            case "line": case "straightConnector1": return new List<Seg> { M(0, 0), L(w, h) };
            case "bentConnector3": return new List<Seg> { M(0, 0), L(w / 2, 0), L(w / 2, h), L(w, h) };
            case "bentConnector2": return new List<Seg> { M(0, 0), L(w, 0), L(w, h) };
            // ponytail: ogni altra geometria predefinita (flowchart, callout, fumetti...) è disegnata come rettangolo → aggiungere il caso in questo switch
            default: return Poly((0, 0), (w, 0), (w, h), (0, h));
        }
    }

    /// <summary>custGeom: moveTo/lnTo/cubicBezTo/quadBezTo/close; arcTo ignorato (ponytail: gli archi dei freeform non sono disegnati).</summary>
    public static List<List<Seg>> Custom(XElement cust, double w, double h)
    {
        var res = new List<List<Seg>>();
        var pl = cust.Element(Ns.A + "pathLst"); if (pl == null) return res;
        foreach (var path in pl.Elements(Ns.A + "path"))
        {
            double pw = (double?)path.Attribute("w") ?? 0, ph = (double?)path.Attribute("h") ?? 0;
            double sx = pw > 0 ? w / pw : 1, sy = ph > 0 ? h / ph : 1;
            (double, double) P(XElement pt) => (((double?)pt.Attribute("x") ?? 0) * sx, ((double?)pt.Attribute("y") ?? 0) * sy);
            var segs = new List<Seg>(); double lx = 0, ly = 0;
            foreach (var c in path.Elements())
            {
                var pts = c.Elements(Ns.A + "pt").Select(P).ToArray();
                switch (c.Name.LocalName)
                {
                    case "moveTo" when pts.Length > 0: segs.Add(M(pts[0].Item1, pts[0].Item2)); (lx, ly) = pts[0]; break;
                    case "lnTo" when pts.Length > 0: segs.Add(L(pts[0].Item1, pts[0].Item2)); (lx, ly) = pts[0]; break;
                    case "cubicBezTo" when pts.Length == 3: segs.Add(C(pts[0].Item1, pts[0].Item2, pts[1].Item1, pts[1].Item2, pts[2].Item1, pts[2].Item2)); (lx, ly) = pts[2]; break;
                    case "quadBezTo" when pts.Length == 2:
                        segs.Add(C(lx + 2.0 / 3 * (pts[0].Item1 - lx), ly + 2.0 / 3 * (pts[0].Item2 - ly), pts[1].Item1 + 2.0 / 3 * (pts[0].Item1 - pts[1].Item1), pts[1].Item2 + 2.0 / 3 * (pts[0].Item2 - pts[1].Item2), pts[1].Item1, pts[1].Item2));
                        (lx, ly) = pts[1]; break;
                    case "close": segs.Add(Z); break;
                }
            }
            if (segs.Count > 0) res.Add(segs);
        }
        return res;
    }

    /// <summary>Applica ribaltamento e rotazione attorno al centro del riquadro e traslazione.</summary>
    public static List<Seg> Place(IEnumerable<Seg> segs, double ox, double oy, double w, double h, double rotDeg, bool flipH, bool flipV)
    {
        double rad = rotDeg * Math.PI / 180, cs = Math.Cos(rad), sn = Math.Sin(rad);
        (double, double) T(double x, double y)
        {
            if (flipH) x = w - x;
            if (flipV) y = h - y;
            double dx = x - w / 2, dy = y - h / 2;
            return (ox + w / 2 + dx * cs - dy * sn, oy + h / 2 + dx * sn + dy * cs);
        }
        var r = new List<Seg>();
        foreach (var s in segs)
        {
            var (a, b) = T(s.X1, s.Y1); var (c, d) = T(s.X2, s.Y2); var (e, f) = T(s.X3, s.Y3);
            r.Add(s.Kind switch { SegKind.Move or SegKind.Line => s with { X1 = a, Y1 = b }, SegKind.Cubic => new Seg(SegKind.Cubic, a, b, c, d, e, f), _ => s });
        }
        return r;
    }
}
