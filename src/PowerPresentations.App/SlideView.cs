using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.App;

/// <summary>Area di modifica: disegna la display list, seleziona, sposta e ridimensiona le forme.</summary>
public sealed class SlideView : FrameworkElement
{
    DisplayList? _dl;
    double _k = 1, _ox, _oy;
    int _sel;
    // trascinamento
    int _dragId; bool _resize; Point _start; RectD _startBox; RectD _liveBox; bool _moved;

    public event Action<int>? Selected;
    public event Action<int, RectD>? BoxChanged;
    public event Action<int>? EditRequested;

    public int SelectedId { get => _sel; set { _sel = value; InvalidateVisual(); } }
    public DisplayList? List { get => _dl; set { _dl = value; InvalidateVisual(); } }

    public SlideView() { Focusable = true; ClipToBounds = true; }

    void Layout_()
    {
        if (_dl == null) return;
        const double m = 28;
        double aw = Math.Max(10, ActualWidth - 2 * m), ah = Math.Max(10, ActualHeight - 2 * m);
        _k = Math.Min(aw / _dl.Width, ah / _dl.Height);
        _ox = (ActualWidth - _dl.Width * _k) / 2; _oy = (ActualHeight - _dl.Height * _k) / 2;
    }

    public Rect ToView(RectD r) { Layout_(); return new Rect(_ox + r.X * _k, _oy + r.Y * _k, r.W * _k, r.H * _k); }
    RectD? BoxOf(int id) => _dl?.Hits.Where(h => h.Id == id).Select(h => (RectD?)h.Box).FirstOrDefault();

    protected override void OnRender(DrawingContext dc)
    {
        dc.DrawRectangle((Brush)Application.Current.Resources["CanvasBrush"], null, new Rect(0, 0, ActualWidth, ActualHeight));
        if (_dl == null) return;
        Layout_();
        dc.PushTransform(new TranslateTransform(_ox, _oy));
        dc.PushTransform(new ScaleTransform(_k, _k));
        SlideDrawing.Draw(dc, _dl);
        dc.Pop(); dc.Pop();
        if (_sel != 0 && (_dragId == _sel && _moved ? _liveBox : BoxOf(_sel)) is { } b)
        {
            var r = ToView(b);
            var accent = (Brush)Application.Current.Resources["AccentBrush"];
            dc.DrawRectangle(null, new Pen(accent, 1.5), r);
            dc.DrawRectangle(accent, null, new Rect(r.Right - 5, r.Bottom - 5, 10, 10));
        }
    }

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        Focus();
        if (_dl == null) return;
        Layout_();
        var p = e.GetPosition(this);
        var pt = new Point((p.X - _ox) / _k, (p.Y - _oy) / _k);
        _resize = false; _moved = false;
        if (_sel != 0 && BoxOf(_sel) is { } sb)
        {
            var r = ToView(sb);
            if (new Rect(r.Right - 8, r.Bottom - 8, 16, 16).Contains(p)) { _resize = true; _dragId = _sel; _start = pt; _startBox = _liveBox = sb; CaptureMouse(); return; }
        }
        var hit = _dl.Hits.LastOrDefault(h => pt.X >= h.Box.X && pt.X <= h.Box.X + h.Box.W && pt.Y >= h.Box.Y && pt.Y <= h.Box.Y + h.Box.H);
        _sel = hit.Id; Selected?.Invoke(_sel);
        if (hit.Id != 0)
        {
            if (e.ClickCount == 2 && hit.HasText) { EditRequested?.Invoke(hit.Id); InvalidateVisual(); return; }
            _dragId = hit.Id; _start = pt; _startBox = _liveBox = hit.Box; CaptureMouse();
        }
        InvalidateVisual();
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        if (!IsMouseCaptured || _dragId == 0) return;
        var p = e.GetPosition(this);
        var pt = new Point((p.X - _ox) / _k, (p.Y - _oy) / _k);
        double dx = pt.X - _start.X, dy = pt.Y - _start.Y;
        if (!_moved && Math.Abs(dx) + Math.Abs(dy) < 3) return;
        _moved = true;
        _liveBox = _resize ? _startBox with { W = Math.Max(10, _startBox.W + dx), H = Math.Max(10, _startBox.H + dy) } : _startBox with { X = _startBox.X + dx, Y = _startBox.Y + dy };
        InvalidateVisual();
    }

    protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
    {
        if (IsMouseCaptured) ReleaseMouseCapture();
        if (_moved && _dragId != 0) BoxChanged?.Invoke(_dragId, _liveBox);
        _moved = false; _dragId = 0; InvalidateVisual();
    }
}
