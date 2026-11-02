using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.App;

/// <summary>Area di modifica: disegna la display list, seleziona, sposta e ridimensiona le forme (8 maniglie).</summary>
public sealed class SlideView : FrameworkElement
{
    DisplayList? _dl;
    double _k = 1, _ox, _oy, _zoom = 1;
    int _sel;
    int _dragId; int _handle = -1; Point _start; RectD _startBox; RectD _liveBox; bool _moved;

    public event Action<int>? Selected;
    public event Action<int, RectD>? BoxChanged;
    public event Action<int, Point>? EditRequested;

    readonly HashSet<int> _multi = new();

    public int SelectedId { get => _sel; set { _sel = value; _multi.Clear(); InvalidateVisual(); } }
    /// <summary>Forme selezionate: una sola (SelectedId) o più (Ctrl+clic, Seleziona tutto).</summary>
    public IReadOnlyList<int> SelectedIds => _multi.Count > 0 ? _multi.ToList() : _sel != 0 ? new[] { _sel } : Array.Empty<int>();
    public void SelectMany(IEnumerable<int> ids)
    {
        var l = ids.Distinct().ToList(); _multi.Clear();
        if (l.Count == 1) _sel = l[0]; else { _sel = 0; foreach (var i in l) _multi.Add(i); }
        InvalidateVisual(); Selected?.Invoke(_sel);
    }
    public DisplayList? List { get => _dl; set { _dl = value; InvalidateVisual(); } }
    /// <summary>1 = la slide sta nella finestra.</summary>
    public double Zoom { get => _zoom; set { _zoom = Math.Clamp(value, 0.25, 2); InvalidateVisual(); } }

    bool _grid, _ruler;
    /// <summary>Griglia (passo 1 cm) e righello (cm) disegnati sopra la slide; non toccano il file.</summary>
    public bool ShowGrid { get => _grid; set { _grid = value; InvalidateVisual(); } }
    public bool ShowRuler { get => _ruler; set { _ruler = value; InvalidateVisual(); } }
    const double Cm = 72 / 2.54;

    public SlideView() { Focusable = true; ClipToBounds = true; }

    void Layout_()
    {
        if (_dl == null) return;
        const double m = 28;
        double aw = Math.Max(10, ActualWidth - 2 * m), ah = Math.Max(10, ActualHeight - 2 * m);
        _k = Math.Min(aw / _dl.Width, ah / _dl.Height) * _zoom;
        _ox = (ActualWidth - _dl.Width * _k) / 2; _oy = (ActualHeight - _dl.Height * _k) / 2;
    }

    public Rect ToView(RectD r) { Layout_(); return new Rect(_ox + r.X * _k, _oy + r.Y * _k, r.W * _k, r.H * _k); }
    RectD? BoxOf(int id) => _dl?.Hits.Where(h => h.Id == id).Select(h => (RectD?)h.Box).FirstOrDefault();

    static Point[] Handles(Rect r) => new[]
    {
        r.TopLeft, new Point(r.X + r.Width / 2, r.Y), r.TopRight, new Point(r.Right, r.Y + r.Height / 2),
        r.BottomRight, new Point(r.X + r.Width / 2, r.Bottom), r.BottomLeft, new Point(r.X, r.Y + r.Height / 2)
    };

    protected override void OnRender(DrawingContext dc)
    {
        dc.DrawRectangle((Brush)Application.Current.Resources["CanvasBrush"], null, new Rect(0, 0, ActualWidth, ActualHeight));
        if (_dl == null) return;
        Layout_();
        dc.PushTransform(new TranslateTransform(_ox, _oy));
        dc.PushTransform(new ScaleTransform(_k, _k));
        SlideDrawing.Draw(dc, _dl);
        dc.Pop(); dc.Pop();
        if (_grid) DrawGrid(dc);
        if (_ruler) DrawRuler(dc);
        foreach (var id in _multi) if (BoxOf(id) is { } mb) dc.DrawRectangle(null, new Pen((Brush)Application.Current.Resources["AccentBrush"], 1.5), ToView(mb));
        if (_sel != 0 && (_dragId == _sel && _moved ? _liveBox : BoxOf(_sel)) is { } b)
        {
            var r = ToView(b);
            var accent = (Brush)Application.Current.Resources["AccentBrush"];
            dc.DrawRectangle(null, new Pen(accent, 1.5), r);
            foreach (var h in Handles(r)) dc.DrawEllipse(Brushes.White, new Pen(accent, 1.5), h, 4.5, 4.5);
        }
    }

    void DrawGrid(DrawingContext dc)
    {
        var pen = new Pen(new SolidColorBrush(Color.FromArgb(70, 128, 128, 128)), 1); pen.Freeze();
        double w = _dl!.Width, h = _dl.Height;
        for (double x = Cm; x < w; x += Cm) dc.DrawLine(pen, new Point(_ox + x * _k, _oy), new Point(_ox + x * _k, _oy + h * _k));
        for (double y = Cm; y < h; y += Cm) dc.DrawLine(pen, new Point(_ox, _oy + y * _k), new Point(_ox + w * _k, _oy + y * _k));
    }

    void DrawRuler(DrawingContext dc)
    {
        const double t = 18;
        var bg = (Brush)Application.Current.Resources["PanelBrush"]; var fg = (Brush)Application.Current.Resources["MutedBrush"];
        var pen = new Pen(fg, 1);
        dc.DrawRectangle(bg, null, new Rect(0, 0, ActualWidth, t)); dc.DrawRectangle(bg, null, new Rect(0, 0, t, ActualHeight));
        double dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        void Label(string s, Point p) => dc.DrawText(new FormattedText(s, System.Globalization.CultureInfo.CurrentUICulture, FlowDirection.LeftToRight, new Typeface("Segoe UI"), 9, fg, dpi), p);
        double w = _dl!.Width / Cm, h = _dl.Height / Cm;
        for (int i = 0; i * 0.5 <= w; i++)
        {
            double x = _ox + i * 0.5 * Cm * _k; bool whole = i % 2 == 0;
            dc.DrawLine(pen, new Point(x, t), new Point(x, t - (whole ? 8 : 4)));
            if (whole && i % 2 == 0) Label((i / 2).ToString(), new Point(x + 2, 1));
        }
        for (int i = 0; i * 0.5 <= h; i++)
        {
            double y = _oy + i * 0.5 * Cm * _k; bool whole = i % 2 == 0;
            dc.DrawLine(pen, new Point(t, y), new Point(t - (whole ? 8 : 4), y));
            if (whole) Label((i / 2).ToString(), new Point(1, y + 1));
        }
    }

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        Focus();
        if (_dl == null) return;
        Layout_();
        var p = e.GetPosition(this);
        var pt = new Point((p.X - _ox) / _k, (p.Y - _oy) / _k);
        _handle = -1; _moved = false;
        if (_sel != 0 && BoxOf(_sel) is { } sb)
        {
            var hs = Handles(ToView(sb));
            for (int i = 0; i < hs.Length; i++)
                if ((hs[i] - p).Length <= 8) { _handle = i; _dragId = _sel; _start = pt; _startBox = _liveBox = sb; CaptureMouse(); return; }
        }
        var hit = _dl.Hits.LastOrDefault(h => pt.X >= h.Box.X && pt.X <= h.Box.X + h.Box.W && pt.Y >= h.Box.Y && pt.Y <= h.Box.Y + h.Box.H);
        if (hit.Id != 0 && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
        {
            // Ctrl+clic: aggiunge o toglie la forma dalla selezione (nessun trascinamento con più forme)
            var cur = SelectedIds.ToList(); if (!cur.Remove(hit.Id)) cur.Add(hit.Id);
            SelectMany(cur); return;
        }
        _multi.Clear(); _sel = hit.Id; Selected?.Invoke(_sel);
        if (hit.Id != 0)
        {
            if (e.ClickCount == 2 && hit.HasText) { EditRequested?.Invoke(hit.Id, pt); InvalidateVisual(); return; }
            _dragId = hit.Id; _start = pt; _startBox = _liveBox = hit.Box; CaptureMouse();
        }
        InvalidateVisual();
    }

    protected override void OnMouseRightButtonDown(MouseButtonEventArgs e)
    {
        // il menu contestuale agisce sulla forma sotto il cursore (se non è già selezionata)
        Focus();
        if (_dl == null) return;
        Layout_(); var p = e.GetPosition(this); var pt = new Point((p.X - _ox) / _k, (p.Y - _oy) / _k);
        var hit = _dl.Hits.LastOrDefault(h => pt.X >= h.Box.X && pt.X <= h.Box.X + h.Box.W && pt.Y >= h.Box.Y && pt.Y <= h.Box.Y + h.Box.H);
        if (!SelectedIds.Contains(hit.Id)) { _multi.Clear(); _sel = hit.Id; Selected?.Invoke(_sel); InvalidateVisual(); }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        if (!IsMouseCaptured || _dragId == 0) return;
        var p = e.GetPosition(this);
        var pt = new Point((p.X - _ox) / _k, (p.Y - _oy) / _k);
        double dx = pt.X - _start.X, dy = pt.Y - _start.Y;
        if (!_moved && Math.Abs(dx) + Math.Abs(dy) < 3) return;
        _moved = true;
        var b = _startBox; double x = b.X, y = b.Y, w = b.W, h = b.H;
        if (_handle < 0) { x += dx; y += dy; }
        else
        {
            // 0 TL, 1 T, 2 TR, 3 R, 4 BR, 5 B, 6 BL, 7 L
            bool left = _handle is 0 or 6 or 7, right = _handle is 2 or 3 or 4, top = _handle is 0 or 1 or 2, bottom = _handle is 4 or 5 or 6;
            if (left) { x += dx; w -= dx; } if (right) w += dx;
            if (top) { y += dy; h -= dy; } if (bottom) h += dy;
            if (w < 10) { if (left) x -= 10 - w; w = 10; }
            if (h < 10) { if (top) y -= 10 - h; h = 10; }
        }
        _liveBox = new RectD(x, y, w, h);
        InvalidateVisual();
    }

    protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
    {
        if (IsMouseCaptured) ReleaseMouseCapture();
        if (_moved && _dragId != 0) BoxChanged?.Invoke(_dragId, _liveBox);
        _moved = false; _dragId = 0; _handle = -1; InvalidateVisual();
    }
}
