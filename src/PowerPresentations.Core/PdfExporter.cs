using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Core;

/// <summary>Esporta le slide in PDF (PDFsharp, MIT) a partire dalla stessa display list usata dallo schermo.</summary>
public static class PdfExporter
{
    sealed class Resolver : IFontResolver
    {
        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic) =>
            new(Rendering.FontFace.Get(familyName, isBold, isItalic).Path);
        public byte[]? GetFont(string faceName) => File.ReadAllBytes(faceName);
    }

    static readonly object Lock = new();
    static void Init() { lock (Lock) { if (GlobalFontSettings.FontResolver == null) GlobalFontSettings.FontResolver = new Resolver(); } }

    static XColor Col(Rgba c) => XColor.FromArgb(c.A, c.R, c.G, c.B);

    /// <param name="includeHidden">PowerPoint di norma non esporta le slide nascoste.</param>
    public static void Export(PptxDocument doc, string path, bool includeHidden = false)
    {
        Init();
        var r = new SlideRenderer(doc);
        using var pdf = new PdfDocument();
        pdf.Info.Title = Path.GetFileNameWithoutExtension(doc.FilePath ?? "presentazione");
        pdf.Info.Creator = "PowerPresentations";
        var slides = doc.Slides;
        for (int i = 0; i < slides.Count; i++)
        {
            if (!includeHidden && doc.IsHidden(i)) continue;
            var dl = r.Render(i);
            var page = pdf.AddPage();
            page.Width = XUnit.FromPoint(dl.Width); page.Height = XUnit.FromPoint(dl.Height);
            using var g = XGraphics.FromPdfPage(page, XGraphicsUnit.Point);
            DrawList(g, dl);
        }
        if (pdf.PageCount == 0) throw new InvalidOperationException("Nessuna slide da esportare.");
        var tmp = path + ".tmp";
        pdf.Save(tmp);
        File.Move(tmp, path, true);
    }

    static XGraphicsPath PathOf(IReadOnlyList<Seg> segs)
    {
        var gp = new XGraphicsPath(); double cx = 0, cy = 0, sx = 0, sy = 0; bool open = false;
        foreach (var s in segs)
            switch (s.Kind)
            {
                case SegKind.Move: if (open) gp.StartFigure(); cx = sx = s.X1; cy = sy = s.Y1; open = true; break;
                case SegKind.Line: gp.AddLine(cx, cy, s.X1, s.Y1); cx = s.X1; cy = s.Y1; break;
                case SegKind.Cubic: gp.AddBezier(cx, cy, s.X1, s.Y1, s.X2, s.Y2, s.X3, s.Y3); cx = s.X3; cy = s.Y3; break;
                case SegKind.Close: gp.CloseFigure(); cx = sx; cy = sy; break;
            }
        return gp;
    }

    static void DrawList(XGraphics g, DisplayList dl)
    {
        var page = new XRect(0, 0, dl.Width, dl.Height);
        XBrush Bg() => dl.Background is SolidFill s ? new XSolidBrush(Col(s.Color)) : dl.Background is GradientFill gf ? Grad(gf, page) : XBrushes.White;
        g.DrawRectangle(Bg(), page);
        g.IntersectClip(page);
        foreach (var it in dl.Items)
            switch (it)
            {
                case PathItem p:
                    {
                        var path = PathOf(p.Segs);
                        XPen? pen = null;
                        if (p.Stroke is { } st)
                        {
                            pen = new XPen(Col(st.Color), st.Width);
                            if (st.Dash != null) { pen.DashStyle = XDashStyle.Custom; pen.DashPattern = st.Dash; }
                        }
                        XBrush? br = p.Fill switch { SolidFill s => new XSolidBrush(Col(s.Color)), GradientFill gf => Grad(gf, Bounds(p.Segs)), _ => null };
                        if (br != null && pen != null) g.DrawPath(pen, br, path);
                        else if (br != null) g.DrawPath(br, path);
                        else if (pen != null) g.DrawPath(pen, path);
                        break;
                    }
                case TextItem t when t.Text.Length > 0:
                    {
                        var font = new XFont(t.Font, t.Size, (t.Bold ? XFontStyleEx.Bold : XFontStyleEx.Regular) | (t.Italic ? XFontStyleEx.Italic : XFontStyleEx.Regular));
                        var state = g.Save();
                        if (t.Rot != 0) g.RotateAtTransform(t.Rot, new XPoint(t.Cx, t.Cy));
                        g.DrawString(t.Text, font, new XSolidBrush(Col(t.Color)), new XPoint(t.X, t.Baseline), XStringFormats.BaseLineLeft);
                        g.Restore(state);
                        break;
                    }
                case ImageItem im: DrawImage(g, im); break;
            }
    }

    static XRect Bounds(IReadOnlyList<Seg> segs)
    {
        double x0 = double.MaxValue, y0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue;
        void P(double x, double y) { x0 = Math.Min(x0, x); y0 = Math.Min(y0, y); x1 = Math.Max(x1, x); y1 = Math.Max(y1, y); }
        foreach (var s in segs)
        {
            if (s.Kind == SegKind.Close) continue;
            P(s.X1, s.Y1); if (s.Kind == SegKind.Cubic) { P(s.X2, s.Y2); P(s.X3, s.Y3); }
        }
        return x0 > x1 ? new XRect(0, 0, 1, 1) : new XRect(x0, y0, Math.Max(1, x1 - x0), Math.Max(1, y1 - y0));
    }

    // ponytail: il PDF usa solo primo e ultimo colore del gradiente
    static XBrush Grad(GradientFill gf, XRect b)
    {
        double a = gf.AngleDeg * Math.PI / 180, dx = Math.Cos(a) / 2, dy = Math.Sin(a) / 2;
        var c = new XPoint(b.X + b.Width / 2, b.Y + b.Height / 2);
        return new XLinearGradientBrush(new XPoint(c.X - dx * b.Width, c.Y - dy * b.Height), new XPoint(c.X + dx * b.Width, c.Y + dy * b.Height), Col(gf.Stops[0].Color), Col(gf.Stops[^1].Color));
    }

    static void DrawImage(XGraphics g, ImageItem im)
    {
        var b = im.Bounds; var state = g.Save();
        try
        {
            double cx = b.X + b.W / 2, cy = b.Y + b.H / 2;
            if (im.Rot != 0) g.RotateAtTransform(im.Rot, new XPoint(cx, cy));
            if (im.FlipH || im.FlipV) g.ScaleAtTransform(im.FlipH ? -1 : 1, im.FlipV ? -1 : 1, new XPoint(cx, cy));
            g.IntersectClip(new XRect(b.X, b.Y, b.W, b.H));
            if (im.Data == null) { g.DrawRectangle(XBrushes.LightGray, new XRect(b.X, b.Y, b.W, b.H)); return; }
            try
            {
                using var xi = XImage.FromStream(new MemoryStream(im.Data));
                double fw = b.W / im.Crop.W, fh = b.H / im.Crop.H;
                g.DrawImage(xi, b.X - im.Crop.X * fw, b.Y - im.Crop.Y * fh, fw, fh);
            }
            catch (Exception e) when (e is NotImplementedException or InvalidOperationException or ArgumentException or NotSupportedException or IOException)
            {
                // ponytail: formati non supportati da PDFsharp (EMF/WMF/SVG) → riquadro grigio
                g.DrawRectangle(XBrushes.LightGray, new XRect(b.X, b.Y, b.W, b.H));
            }
        }
        finally { g.Restore(state); }
    }
}
