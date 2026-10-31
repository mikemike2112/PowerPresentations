using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Geometry = System.Windows.Media.Geometry;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.App;

/// <summary>Disegna una display list con WPF. Coordinate in punti: il chiamante imposta la scala.</summary>
public static class SlideDrawing
{
    static readonly ConditionalWeakTable<byte[], BitmapSource?> Images = new();

    static Color C(Rgba c) => Color.FromArgb(c.A, c.R, c.G, c.B);

    static Brush? BrushOf(Fill? f)
    {
        switch (f)
        {
            case SolidFill s: { var b = new SolidColorBrush(C(s.Color)); b.Freeze(); return b; }
            case GradientFill g:
                {
                    double a = g.AngleDeg * Math.PI / 180;
                    var lb = new LinearGradientBrush
                    {
                        StartPoint = new Point(0.5 - Math.Cos(a) / 2, 0.5 - Math.Sin(a) / 2),
                        EndPoint = new Point(0.5 + Math.Cos(a) / 2, 0.5 + Math.Sin(a) / 2)
                    };
                    foreach (var s in g.Stops) lb.GradientStops.Add(new GradientStop(C(s.Color), s.Pos));
                    lb.Freeze(); return lb;
                }
        }
        return null;
    }

    static Geometry GeometryOf(IReadOnlyList<Seg> segs, bool filled)
    {
        var g = new StreamGeometry();
        using (var c = g.Open())
        {
            bool open = false;
            foreach (var s in segs)
                switch (s.Kind)
                {
                    case SegKind.Move: c.BeginFigure(new Point(s.X1, s.Y1), filled, false); open = true; break;
                    case SegKind.Line: if (open) c.LineTo(new Point(s.X1, s.Y1), true, false); break;
                    case SegKind.Cubic: if (open) c.BezierTo(new Point(s.X1, s.Y1), new Point(s.X2, s.Y2), new Point(s.X3, s.Y3), true, false); break;
                    case SegKind.Close: break; // chiusura gestita sotto
                }
        }
        // ponytail: tracciati chiusi → figura chiusa se contiene Close
        if (segs.Any(s => s.Kind == SegKind.Close))
        {
            var g2 = new StreamGeometry();
            using var c = g2.Open();
            bool open = false;
            foreach (var s in segs)
                switch (s.Kind)
                {
                    case SegKind.Move: c.BeginFigure(new Point(s.X1, s.Y1), filled, true); open = true; break;
                    case SegKind.Line: if (open) c.LineTo(new Point(s.X1, s.Y1), true, false); break;
                    case SegKind.Cubic: if (open) c.BezierTo(new Point(s.X1, s.Y1), new Point(s.X2, s.Y2), new Point(s.X3, s.Y3), true, false); break;
                }
            c.Close(); g2.Freeze(); return g2;
        }
        g.Freeze(); return g;
    }

    public static void Draw(DrawingContext dc, DisplayList dl)
    {
        var bg = BrushOf(dl.Background) ?? Brushes.White;
        dc.DrawRectangle(bg, null, new Rect(0, 0, dl.Width, dl.Height));
        dc.PushClip(new RectangleGeometry(new Rect(0, 0, dl.Width, dl.Height)));
        foreach (var it in dl.Items)
            switch (it)
            {
                case PathItem p:
                    {
                        Pen? pen = null;
                        if (p.Stroke is { } st)
                        {
                            pen = new Pen(new SolidColorBrush(C(st.Color)), st.Width) { LineJoin = PenLineJoin.Miter };
                            if (st.Dash != null) pen.DashStyle = new DashStyle(st.Dash, 0);
                        }
                        dc.DrawGeometry(BrushOf(p.Fill), pen, GeometryOf(p.Segs, p.Fill != null));
                        break;
                    }
                case TextItem t: DrawText(dc, t); break;
                case ImageItem im: DrawImage(dc, im); break;
            }
        dc.Pop();
    }

    static readonly Dictionary<string, GlyphTypeface?> Gtf = new();

    static GlyphTypeface? Gt(string path)
    {
        lock (Gtf)
        {
            if (!Gtf.TryGetValue(path, out var g))
            {
                try { g = new GlyphTypeface(new Uri(path)); } catch (Exception e) when (e is IOException or NotSupportedException or ArgumentException or UriFormatException) { g = null; }
                Gtf[path] = g;
            }
            return g;
        }
    }

    /// <summary>Il testo si disegna con GlyphRun dal file del font usato dal layout: stesse larghezze misurate dal motore di impaginazione.</summary>
    static void DrawText(DrawingContext dc, TextItem t)
    {
        if (t.Text.Length == 0) return;
        var brush = new SolidColorBrush(C(t.Color));
        bool rot = t.Rot != 0;
        if (rot) dc.PushTransform(new RotateTransform(t.Rot, t.Cx, t.Cy));
        try
        {
            var gt = Gt(FontFace.Get(t.Font, t.Bold, t.Italic).Path);
            var idx = new List<ushort>(); var adv = new List<double>(); bool ok = gt != null;
            if (gt != null)
                for (int i = 0; i < t.Text.Length && ok; i++)
                {
                    int cp = char.ConvertToUtf32(t.Text, i); if (char.IsHighSurrogate(t.Text[i])) i++;
                    if (gt.CharacterToGlyphMap.TryGetValue(cp, out var g)) { idx.Add(g); adv.Add(gt.AdvanceWidths[g] * t.Size); } else ok = false;
                }
            if (ok && gt != null)
                dc.DrawGlyphRun(brush, new GlyphRun(gt, 0, false, t.Size, 1.0f, idx, new Point(t.X, t.Baseline), adv, null, null, null, null, null, null));
            else
            {
                // glifo assente nel font (simboli, emoji...): ripiego su FormattedText con la famiglia
                var tf = new Typeface(new FontFamily(t.Font), t.Italic ? FontStyles.Italic : FontStyles.Normal, t.Bold ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal);
                var ft = new FormattedText(t.Text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight, tf, t.Size, brush, 1.0);
                dc.DrawText(ft, new Point(t.X, t.Baseline - ft.Baseline));
            }
        }
        finally { if (rot) dc.Pop(); }
    }

    static BitmapSource? Decode(byte[] data)
    {
        return Images.GetValue(data, d =>
        {
            try
            {
                var bf = BitmapFrame.Create(new MemoryStream(d), BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
                bf.Freeze(); return bf;
            }
            catch (Exception e) when (e is NotSupportedException or FileFormatException or IOException or InvalidOperationException) { return null; }
        });
    }

    static void DrawImage(DrawingContext dc, ImageItem im)
    {
        var b = im.Bounds; var bmp = im.Data == null ? null : Decode(im.Data);
        double cx = b.X + b.W / 2, cy = b.Y + b.H / 2;
        int pushed = 0;
        if (im.Rot != 0) { dc.PushTransform(new RotateTransform(im.Rot, cx, cy)); pushed++; }
        if (im.FlipH || im.FlipV) { dc.PushTransform(new ScaleTransform(im.FlipH ? -1 : 1, im.FlipV ? -1 : 1, cx, cy)); pushed++; }
        if (bmp == null)
        {
            // ponytail: formati non decodificabili da WPF (EMF/WMF, SVG senza PNG) → riquadro grigio
            dc.DrawRectangle(new SolidColorBrush(Color.FromRgb(225, 225, 225)), new Pen(Brushes.Gray, 0.75), new Rect(b.X, b.Y, b.W, b.H));
        }
        else
        {
            dc.PushClip(new RectangleGeometry(new Rect(b.X, b.Y, b.W, b.H)));
            double fw = b.W / im.Crop.W, fh = b.H / im.Crop.H;
            dc.DrawImage(bmp, new Rect(b.X - im.Crop.X * fw, b.Y - im.Crop.Y * fh, fw, fh));
            dc.Pop();
        }
        for (; pushed > 0; pushed--) dc.Pop();
    }

    /// <summary>Rasterizza la display list a una larghezza in pixel.</summary>
    public static BitmapSource ToBitmap(DisplayList dl, int widthPx)
    {
        double k = widthPx / dl.Width;
        int h = (int)Math.Round(dl.Height * k);
        var dv = new DrawingVisual();
        using (var dc = dv.RenderOpen())
        {
            dc.PushTransform(new ScaleTransform(k, k));
            Draw(dc, dl);
            dc.Pop();
        }
        var rtb = new RenderTargetBitmap(widthPx, h, 96, 96, PixelFormats.Pbgra32);
        rtb.Render(dv);
        rtb.Freeze();
        return rtb;
    }

    public static void SavePng(BitmapSource bmp, string path)
    {
        var enc = new PngBitmapEncoder(); enc.Frames.Add(BitmapFrame.Create(bmp));
        using var fs = File.Create(path); enc.Save(fs);
    }
}
