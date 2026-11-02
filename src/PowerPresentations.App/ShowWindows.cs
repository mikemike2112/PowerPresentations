using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.App;

/// <summary>Stato della presentazione condiviso fra schermo pieno e vista relatore.</summary>
public sealed class ShowController
{
    public PptxDocument Doc { get; }
    public ShowSettings Settings { get; }
    /// <summary>Ciclo continuo (anche in modalità postazione informatica).</summary>
    public bool Loop => Settings.Loop || Settings.Mode == "kiosk";
    public SlideRenderer Renderer { get; }
    public List<int> Order { get; }
    public int Pos { get; private set; }
    public bool Black { get; private set; }
    public DateTime Started { get; } = DateTime.Now;
    public event Action<int>? Changed; // argomento: posizione precedente
    /// <summary>Esc: chiude insieme presentazione e vista relatore.</summary>
    public event Action? EndRequested;
    public int Current => Order[Pos];
    public int? NextIndex => Pos + 1 < Order.Count ? Order[Pos + 1] : null;

    public ShowController(PptxDocument doc, SlideRenderer r, int start)
    {
        Doc = doc; Renderer = r;
        Settings = doc.GetShowSettings();
        // intervallo di Imposta presentazione: vale solo se si parte da dentro l'intervallo
        bool inRange = Settings.From is { } f && Settings.To is { } t && start >= f - 1 && start <= t - 1;
        Order = Enumerable.Range(0, doc.Slides.Count).Where(i => (!doc.IsHidden(i) || i == start) && (!inRange || (i >= Settings.From!.Value - 1 && i <= Settings.To!.Value - 1))).ToList();
        Pos = Math.Max(0, Order.IndexOf(start));
    }

    public bool Go(int pos)
    {
        pos = Math.Clamp(pos, 0, Order.Count - 1);
        if (pos == Pos && !Black) return false;
        var old = Pos; Pos = pos; Black = false; Changed?.Invoke(old); return true;
    }
    public bool Next() => Pos + 1 < Order.Count ? Go(Pos + 1) : Loop && Order.Count > 1 && Go(0);   // ponytail: kiosk = solo ciclo, i clic avanzano comunque
    public bool Prev() => Pos > 0 && Go(Pos - 1);
    public void ToggleBlack() { Black = !Black; Changed?.Invoke(Pos); }

    /// <summary>Tasti comuni a presentazione e vista relatore. Ritorna true se gestito.</summary>
    public bool Key(Key k, Action close)
    {
        switch (k)
        {
            case System.Windows.Input.Key.Right: case System.Windows.Input.Key.Down: case System.Windows.Input.Key.Space: case System.Windows.Input.Key.Enter: case System.Windows.Input.Key.PageDown: case System.Windows.Input.Key.N: Next(); return true;
            case System.Windows.Input.Key.Left: case System.Windows.Input.Key.Up: case System.Windows.Input.Key.Back: case System.Windows.Input.Key.PageUp: case System.Windows.Input.Key.P: Prev(); return true;
            case System.Windows.Input.Key.Home: Go(0); return true;
            case System.Windows.Input.Key.End: Go(Order.Count - 1); return true;
            case System.Windows.Input.Key.B: case System.Windows.Input.Key.OemPeriod: ToggleBlack(); return true;
            case System.Windows.Input.Key.Escape: if (EndRequested != null) EndRequested.Invoke(); else close(); return true;
        }
        return false;
    }
}

public sealed class PresentationWindow : Window
{
    readonly ShowController _c;
    readonly Image _back = new() { Stretch = Stretch.Uniform }, _front = new() { Stretch = Stretch.Uniform };
    readonly Grid _root = new() { Background = Brushes.Black };
    BitmapSource? _cur;
    public bool Animate { get; set; } = true;

    public PresentationWindow(ShowController c)
    {
        _c = c;
        WindowStyle = WindowStyle.None; ResizeMode = ResizeMode.NoResize; Background = Brushes.Black; Cursor = Cursors.None; Title = "Presentazione";
        _root.Children.Add(_back); _root.Children.Add(_front); Content = _root;
        _root.ClipToBounds = true;
        Loaded += (_, _) => { _cur = Render(_c.Current); _back.Source = _cur; };
        KeyDown += (_, e) => e.Handled = _c.Key(e.Key, Close);
        MouseLeftButtonDown += (_, _) => _c.Next();
        MouseRightButtonDown += (_, _) => _c.Prev();
        _c.Changed += OnChanged; _c.EndRequested += Close;
        Closed += (_, _) => { _c.Changed -= OnChanged; _c.EndRequested -= Close; };
    }

    BitmapSource Render(int index)
    {
        var dl = _c.Renderer.Render(index);
        var src = PresentationSource.FromVisual(this); double dpi = src?.CompositionTarget?.TransformToDevice.M11 ?? 1;
        int w = (int)Math.Clamp((ActualWidth > 0 ? ActualWidth : SystemParameters.PrimaryScreenWidth) * dpi, 640, 3840);
        return SlideDrawing.ToBitmap(dl, w);
    }

    void OnChanged(int oldPos)
    {
        if (_c.Black) { _back.Source = null; _front.Source = null; return; }
        var next = Render(_c.Current);
        var tr = _c.Doc.GetTransition(_c.Current);
        if (!Animate || oldPos == _c.Pos || tr.Kind == "none" || tr.Kind == "cut" || _cur == null) { _back.Source = next; _front.Source = null; _cur = next; return; }
        Transition(tr, _cur, next);
        _cur = next;
    }

    void Transition(TransitionInfo tr, BitmapSource from, BitmapSource to)
    {
        double w = ActualWidth, h = ActualHeight;
        var secs = tr.Speed switch { "slow" => 1.0, "fast" => 0.4, _ => 0.7 };
        var dur = new Duration(TimeSpan.FromSeconds(secs));
        _back.Source = from; _front.Source = to;
        _back.RenderTransform = null; _front.RenderTransform = null; _front.Clip = null; _front.Opacity = 1; _back.Opacity = 1;
        Panel.SetZIndex(_front, 1); Panel.SetZIndex(_back, 0);
        void Done() { _back.Source = to; _front.Source = null; _front.Clip = null; _front.RenderTransform = null; _back.RenderTransform = null; _front.Opacity = 1; }
        // dir: direzione del movimento (l = verso sinistra, la nuova slide arriva da destra)
        (double fx, double fy) = tr.Dir switch { "r" => (-w, 0.0), "u" => (0.0, h), "d" => (0.0, -h), _ => (w, 0.0) };
        AnimationTimeline? last = null;
        switch (tr.Kind)
        {
            case "push":
                {
                    var tf = new TranslateTransform(fx, fy); var tb = new TranslateTransform(); _front.RenderTransform = tf; _back.RenderTransform = tb;
                    tf.BeginAnimation(TranslateTransform.XProperty, new DoubleAnimation(fx, 0, dur)); tf.BeginAnimation(TranslateTransform.YProperty, new DoubleAnimation(fy, 0, dur));
                    tb.BeginAnimation(TranslateTransform.XProperty, new DoubleAnimation(0, -fx, dur)); last = new DoubleAnimation(0, -fy, dur); tb.BeginAnimation(TranslateTransform.YProperty, (DoubleAnimation)last);
                    break;
                }
            case "cover":
                {
                    var tf = new TranslateTransform(fx, fy); _front.RenderTransform = tf;
                    tf.BeginAnimation(TranslateTransform.XProperty, new DoubleAnimation(fx, 0, dur)); last = new DoubleAnimation(fy, 0, dur); tf.BeginAnimation(TranslateTransform.YProperty, (DoubleAnimation)last);
                    break;
                }
            case "pull":
                {
                    Panel.SetZIndex(_front, 0); Panel.SetZIndex(_back, 1); _front.Source = to; _back.Source = from;
                    var tb = new TranslateTransform(); _back.RenderTransform = tb;
                    tb.BeginAnimation(TranslateTransform.XProperty, new DoubleAnimation(0, -fx, dur)); last = new DoubleAnimation(0, -fy, dur); tb.BeginAnimation(TranslateTransform.YProperty, (DoubleAnimation)last);
                    break;
                }
            case "wipe":
                {
                    var clip = new RectangleGeometry(); _front.Clip = clip;
                    Rect a = tr.Dir switch { "r" => new Rect(0, 0, 0, h), "u" => new Rect(0, h, w, 0), "d" => new Rect(0, 0, w, 0), _ => new Rect(w, 0, 0, h) };
                    last = new RectAnimation(a, new Rect(0, 0, w, h), dur); clip.BeginAnimation(RectangleGeometry.RectProperty, (RectAnimation)last);
                    break;
                }
            default: // fade, split e le altre: dissolvenza (ponytail)
                last = new DoubleAnimation(0, 1, dur); _front.BeginAnimation(OpacityProperty, (DoubleAnimation)last); break;
        }
        last!.Completed += (_, _) => { _front.BeginAnimation(OpacityProperty, null); Done(); };
    }

    public void ShowFullScreen(Rect? area)
    {
        if (area is { } a) { WindowStartupLocation = WindowStartupLocation.Manual; Left = a.Left; Top = a.Top; Width = a.Width; Height = a.Height; WindowState = WindowState.Normal; Topmost = true; }
        else { WindowState = WindowState.Maximized; Topmost = true; }
        Show();
    }
}

public sealed class PresenterWindow : Window
{
    readonly ShowController _c;
    readonly Image _cur = new() { Stretch = Stretch.Uniform }, _next = new() { Stretch = Stretch.Uniform };
    readonly TextBlock _notes = new() { TextWrapping = TextWrapping.Wrap, FontSize = 20 }, _info = new() { FontSize = 16 }, _timer = new() { FontSize = 28, FontWeight = FontWeights.SemiBold };
    readonly DispatcherTimer _t = new() { Interval = TimeSpan.FromSeconds(1) };

    public PresenterWindow(ShowController c)
    {
        _c = c; Title = "Vista relatore — PowerPresentations"; Width = 1200; Height = 720;
        Background = (Brush)Application.Current.Resources["BgBrush"]; Foreground = (Brush)Application.Current.Resources["FgBrush"];
        var g = new Grid { Margin = new Thickness(0), Background = (Brush)Application.Current.Resources["BgBrush"] };
        g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(3, GridUnitType.Star) });
        g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
        var left = new Border { Margin = new Thickness(12, 12, 24, 12), BorderBrush = (Brush)Application.Current.Resources["AccentBrush"], BorderThickness = new Thickness(2), Child = _cur, Background = Brushes.Black };
        g.Children.Add(left);
        var right = new Grid { Margin = new Thickness(0, 12, 12, 12) }; right.RowDefinitions.Add(new RowDefinition { Height = new GridLength(2, GridUnitType.Star) }); right.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); right.RowDefinitions.Add(new RowDefinition { Height = new GridLength(3, GridUnitType.Star) });
        right.Children.Add(new Border { BorderBrush = (Brush)Application.Current.Resources["LineBrush"], BorderThickness = new Thickness(1), Child = _next, Background = Brushes.Black });
        var bar = new StackPanel { Margin = new Thickness(0, 8, 0, 8) }; Grid.SetRow(bar, 1); bar.Children.Add(_timer); bar.Children.Add(_info);
        var btns = new StackPanel { Orientation = Orientation.Horizontal };
        var prev = new Button { Content = "◀ Indietro" }; prev.Click += (_, _) => _c.Prev();
        var next = new Button { Content = "Avanti ▶" }; next.Click += (_, _) => _c.Next();
        btns.Children.Add(prev); btns.Children.Add(next); bar.Children.Add(btns);
        right.Children.Add(bar);
        var sv = new ScrollViewer { Content = _notes, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }; Grid.SetRow(sv, 2); right.Children.Add(sv);
        Grid.SetColumn(right, 1); g.Children.Add(right); Content = g;
        KeyDown += (_, e) => e.Handled = _c.Key(e.Key, Close);
        _c.Changed += OnChanged; _c.EndRequested += Close; Closed += (_, _) => { _c.Changed -= OnChanged; _c.EndRequested -= Close; _t.Stop(); };
        _t.Tick += (_, _) => _timer.Text = (DateTime.Now - _c.Started).ToString(@"hh\:mm\:ss");
        _t.Start(); _timer.Text = "00:00:00";
        Loaded += (_, _) => Refresh();
    }

    void OnChanged(int _) => Refresh();

    void Refresh()
    {
        _cur.Source = SlideDrawing.ToBitmap(_c.Renderer.Render(_c.Current), 1280);
        _next.Source = _c.NextIndex is { } n ? SlideDrawing.ToBitmap(_c.Renderer.Render(n), 640) : null;
        var notes = _c.Doc.GetNotes(_c.Current);
        _notes.Text = notes.Length > 0 ? notes : "(nessuna nota)";
        _notes.Opacity = notes.Length > 0 ? 1 : 0.5;
        _info.Text = $"Slide {_c.Pos + 1} di {_c.Order.Count}" + (_c.NextIndex == null ? " — ultima" : "");
    }
}

/// <summary>Finestra di scelta da un elenco (layout, forme).</summary>
public static class ListPicker
{
    public static int Pick(Window owner, string title, IEnumerable<string> items)
    {
        var lb = new ListBox { ItemsSource = items.ToList(), Background = (Brush)Application.Current.Resources["PanelBrush"], Foreground = (Brush)Application.Current.Resources["FgBrush"], BorderThickness = new Thickness(0), Margin = new Thickness(8) };
        lb.ItemContainerStyle = null;
        int result = -1;
        var w = new Window { Title = title, Owner = owner, Width = 340, Height = 420, WindowStartupLocation = WindowStartupLocation.CenterOwner, Background = (Brush)Application.Current.Resources["BgBrush"], Foreground = (Brush)Application.Current.Resources["FgBrush"], Content = lb, ShowInTaskbar = false };
        lb.MouseDoubleClick += (_, _) => { if (lb.SelectedIndex >= 0) { result = lb.SelectedIndex; w.Close(); } };
        lb.KeyDown += (_, e) => { if (e.Key == Key.Enter && lb.SelectedIndex >= 0) { result = lb.SelectedIndex; w.Close(); } else if (e.Key == Key.Escape) w.Close(); };
        w.Loaded += (_, _) => { lb.SelectedIndex = 0; lb.Focus(); };
        w.ShowDialog();
        return result;
    }
}
