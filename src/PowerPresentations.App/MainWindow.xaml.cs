using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.App;

public sealed class ThumbItem : INotifyPropertyChanged
{
    ImageSource? _img;
    public int Number { get; set; }
    public double Opacity { get; set; } = 1;
    public double Width { get; set; } = 170;
    public ImageSource? Image { get => _img; set { _img = value; PropertyChanged?.Invoke(this, new(nameof(Image))); } }
    public event PropertyChangedEventHandler? PropertyChanged;
}

public partial class MainWindow : Window
{
    PptxDocument? _doc;
    SlideRenderer? _r;
    int _cur;
    bool _loading;
    int _editId;
    bool _cancelEdit;
    bool _dark = true;
    bool _sorter;
    readonly ObservableCollection<ThumbItem> _thumbs = new();
    Point _dragStart; int _dragIndex = -1;

    static readonly (string Name, string Prst)[] Shapes =
    {
        ("Rettangolo", "rect"), ("Rettangolo arrotondato", "roundRect"), ("Ellisse", "ellipse"), ("Triangolo", "triangle"), ("Triangolo rettangolo", "rtTriangle"),
        ("Rombo", "diamond"), ("Parallelogramma", "parallelogram"), ("Trapezio", "trapezoid"), ("Esagono", "hexagon"), ("Ottagono", "octagon"),
        ("Croce", "plus"), ("Freccia destra", "rightArrow"), ("Freccia sinistra", "leftArrow"), ("Freccia su", "upArrow"), ("Freccia giù", "downArrow"),
        ("Stella a 5 punte", "star5"), ("Cheuron", "chevron"), ("Pentagono (freccia)", "homePlate")
    };

    public MainWindow()
    {
        InitializeComponent();
        Thumbs.ItemsSource = _thumbs;
        View.Selected += id => { UpdateStatus(); };
        View.BoxChanged += OnBoxChanged;
        View.EditRequested += BeginEdit;
        PreviewKeyDown += OnWindowKey;
        Closing += OnClosing;
        Thumbs.MouseDoubleClick += OnThumbDoubleClick;
        Drop += (_, e) => { if (e.Data.GetData(DataFormats.FileDrop) is string[] f && f.Length > 0 && f[0].EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) && Confirm()) OpenFile(f[0]); };
        NewDoc();
    }

    // ---------- documento ----------
    void NewDoc()
    {
        try { SetDoc(PptxDocument.New()); }
        catch (Exception e) when (e is InvalidOperationException or PptxFormatException) { MessageBox.Show(this, e.Message, "Nuova presentazione", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    public void OpenFile(string path)
    {
        try { SetDoc(PptxDocument.Open(path)); }
        catch (PptxFormatException e) { MessageBox.Show(this, e.Message, "Impossibile aprire il file", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    void SetDoc(PptxDocument d)
    {
        _doc?.Dispose(); _doc = d; _r = new SlideRenderer(d); _cur = 0; _editId = 0; Editor.Visibility = Visibility.Collapsed;
        RebuildThumbs(); ShowSlide(0); UpdateTitle();
    }

    bool Confirm()
    {
        if (_doc is not { Dirty: true }) return true;
        var r = MessageBox.Show(this, "Salvare le modifiche alla presentazione?", "PowerPresentations", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (r == MessageBoxResult.Cancel) return false;
        return r == MessageBoxResult.No || DoSave(false);
    }

    void OnClosing(object? s, CancelEventArgs e) { CommitEdit(); SaveNotes(); if (!Confirm()) e.Cancel = true; else _doc?.Dispose(); }

    void UpdateTitle() => Title = $"{(_doc?.FilePath is { } p ? Path.GetFileName(p) : "Senza titolo")}{(_doc?.Dirty == true ? " *" : "")} — PowerPresentations";

    void UpdateStatus() =>
        Status.Text = _doc == null ? "" : $"Slide {_cur + 1} di {_doc.Slides.Count}" + (View.SelectedId != 0 ? $"   ·   forma {View.SelectedId}" : "") + (_doc.IsHidden(_cur) ? "   ·   nascosta" : "");

    // ---------- miniature ----------
    ThumbItem MakeThumb(int i)
    {
        var bmp = SlideDrawing.ToBitmap(_r!.Render(i), 320);
        return new ThumbItem { Number = i + 1, Image = bmp, Opacity = _doc!.IsHidden(i) ? 0.4 : 1, Width = _sorter ? 300 : 170 };
    }

    void RebuildThumbs()
    {
        _loading = true; _thumbs.Clear();
        for (int i = 0; i < _doc!.Slides.Count; i++) _thumbs.Add(MakeThumb(i));
        Thumbs.SelectedIndex = Math.Min(_cur, _thumbs.Count - 1);
        _loading = false;
    }

    void RefreshCurrent()
    {
        if (_doc == null) return;
        View.List = _r!.Render(_cur);
        if (_cur < _thumbs.Count) _thumbs[_cur] = MakeThumb(_cur);
        _loading = true; Thumbs.SelectedIndex = _cur; _loading = false;
        UpdateTitle(); UpdateStatus();
    }

    void ShowSlide(int i)
    {
        if (_doc == null || _doc.Slides.Count == 0) return;
        CommitEdit();
        _cur = Math.Clamp(i, 0, _doc.Slides.Count - 1);
        View.SelectedId = 0; View.List = _r!.Render(_cur);
        Notes.Text = _doc.GetNotes(_cur);
        UpdateTransUi(); UpdateStatus();
        if (Thumbs.SelectedIndex != _cur) { _loading = true; Thumbs.SelectedIndex = _cur; _loading = false; }
    }

    void OnThumbSelected(object s, SelectionChangedEventArgs e) { if (!_loading && Thumbs.SelectedIndex >= 0) { SaveNotes(); ShowSlide(Thumbs.SelectedIndex); } }

    // riordino con drag & drop: modifica solo sldIdLst
    void OnThumbMouseDown(object s, MouseButtonEventArgs e)
    {
        _dragStart = e.GetPosition(Thumbs);
        var c = ItemsControl.ContainerFromElement(Thumbs, (DependencyObject)e.OriginalSource) as ListBoxItem;
        _dragIndex = c == null ? -1 : Thumbs.ItemContainerGenerator.IndexFromContainer(c);
    }

    void OnThumbMouseMove(object s, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _dragIndex < 0) return;
        var d = e.GetPosition(Thumbs) - _dragStart;
        if (Math.Abs(d.Y) < SystemParameters.MinimumVerticalDragDistance && Math.Abs(d.X) < SystemParameters.MinimumHorizontalDragDistance) return;
        int idx = _dragIndex; _dragIndex = -1;
        DragDrop.DoDragDrop(Thumbs, new DataObject("pp-slide", idx), DragDropEffects.Move);
    }

    void OnThumbDragOver(object s, DragEventArgs e) { e.Effects = e.Data.GetDataPresent("pp-slide") ? DragDropEffects.Move : DragDropEffects.None; e.Handled = true; }

    void OnThumbDrop(object s, DragEventArgs e)
    {
        if (_doc == null || !e.Data.GetDataPresent("pp-slide")) return;
        int from = (int)e.Data.GetData("pp-slide")!;
        int to = _thumbs.Count - 1;
        if (e.OriginalSource is DependencyObject o && ItemsControl.ContainerFromElement(Thumbs, o) is ListBoxItem c) to = Thumbs.ItemContainerGenerator.IndexFromContainer(c);
        if (to < 0 || to == from) return;
        SaveNotes();
        _doc.MoveSlide(from, to); _cur = to; RebuildThumbs(); ShowSlide(to); UpdateTitle();
    }

    void OnThumbDoubleClick(object s, MouseButtonEventArgs e) { if (_sorter && Thumbs.SelectedIndex >= 0) { int i = Thumbs.SelectedIndex; SetSorter(false); ShowSlide(i); } }

    void OnThumbKey(object s, KeyEventArgs e) { if (e.Key == Key.Delete) { OnDeleteSlide(s, e); e.Handled = true; } }

    // vista Sequenza diapositive: le miniature occupano tutta l'area (griglia), il riordino con drag & drop funziona uguale
    void SetSorter(bool on)
    {
        if (_doc == null) return;
        CommitEdit(); SaveNotes(); _sorter = on;
        Grid.SetColumnSpan(Thumbs, on ? 2 : 1); Grid.SetRowSpan(Thumbs, 2);
        EditArea.Visibility = NotesArea.Visibility = on ? Visibility.Collapsed : Visibility.Visible;
        var panel = on ? new FrameworkElementFactory(typeof(WrapPanel)) : new FrameworkElementFactory(typeof(VirtualizingStackPanel));
        Thumbs.ItemsPanel = new ItemsPanelTemplate(panel);
        RebuildThumbs();
    }

    void OnSorterView(object s, RoutedEventArgs e) => SetSorter(true);
    void OnNormalView(object s, RoutedEventArgs e) => SetSorter(false);

    // ---------- file ----------
    void OnNew(object s, RoutedEventArgs e) { if (Confirm()) NewDoc(); }

    void OnOpen(object s, RoutedEventArgs e)
    {
        if (!Confirm()) return;
        var dlg = new OpenFileDialog { Filter = "Presentazioni (*.pptx)|*.pptx", Title = "Apri presentazione" };
        if (dlg.ShowDialog(this) == true) OpenFile(dlg.FileName);
    }

    void OnSave(object s, RoutedEventArgs e) => DoSave(false);
    void OnSaveAs(object s, RoutedEventArgs e) => DoSave(true);

    bool DoSave(bool saveAs)
    {
        if (_doc == null) return false;
        CommitEdit(); SaveNotes();
        string? path = _doc.FilePath;
        if (saveAs || path == null)
        {
            var dlg = new SaveFileDialog { Filter = "Presentazione (*.pptx)|*.pptx", FileName = path != null ? Path.GetFileName(path) : "Presentazione.pptx", DefaultExt = ".pptx" };
            if (dlg.ShowDialog(this) != true) return false;
            path = dlg.FileName;
        }
        try { _doc.SaveAs(path); UpdateTitle(); return true; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
        {
            MessageBox.Show(this, "Salvataggio non riuscito, il file originale non è stato modificato.\n" + ex.Message, "Salva", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    void OnExportPdf(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        CommitEdit();
        var dlg = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", FileName = Path.ChangeExtension(_doc.FilePath != null ? Path.GetFileName(_doc.FilePath) : "Presentazione", ".pdf"), DefaultExt = ".pdf" };
        if (dlg.ShowDialog(this) != true) return;
        try { PdfExporter.Export(_doc, dlg.FileName); Status.Text = "PDF esportato: " + dlg.FileName; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
        { MessageBox.Show(this, "Esportazione PDF non riuscita: " + ex.Message, "Esporta PDF", MessageBoxButton.OK, MessageBoxImage.Error); }
    }

    // ---------- slide ----------
    void OnNewSlide(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        var lays = _doc.Layouts;
        int k = ListPicker.Pick(this, "Layout della nuova slide", lays.Select(l => l.Name));
        if (k < 0) return;
        SaveNotes();
        _cur = _doc.AddSlide(lays[k], _cur); RebuildThumbs(); ShowSlide(_cur); UpdateTitle();
    }

    void OnDuplicate(object s, RoutedEventArgs e) { if (_doc == null) return; SaveNotes(); _cur = _doc.DuplicateSlide(_cur); RebuildThumbs(); ShowSlide(_cur); UpdateTitle(); }

    void OnDeleteSlide(object s, RoutedEventArgs e)
    {
        if (_doc == null || _doc.Slides.Count <= 1) { Status.Text = "Non si può eliminare l'ultima slide."; return; }
        _doc.DeleteSlide(_cur); _cur = Math.Min(_cur, _doc.Slides.Count - 1); RebuildThumbs(); ShowSlide(_cur); UpdateTitle();
    }

    void OnToggleHidden(object s, RoutedEventArgs e) { if (_doc == null) return; _doc.SetHidden(_cur, !_doc.IsHidden(_cur)); RefreshCurrent(); UpdateStatus(); }

    void SaveNotes()
    {
        if (_doc == null || _cur >= _doc.Slides.Count) return;
        if (Notes.Text != _doc.GetNotes(_cur)) { _doc.SetNotes(_cur, Notes.Text); UpdateTitle(); }
    }

    void OnNotesLost(object s, RoutedEventArgs e) => SaveNotes();

    // ---------- modifica testo e oggetti ----------
    void BeginEdit(int id)
    {
        if (_doc == null) return;
        var hit = View.List!.Hits.FirstOrDefault(h => h.Id == id); if (hit.Id == 0) return;
        _editId = id; _cancelEdit = false;
        var r = View.ToView(hit.Box);
        Editor.Margin = new Thickness(r.X, r.Y, 0, 0); Editor.Width = Math.Max(80, r.Width); Editor.MinHeight = Math.Max(28, r.Height);
        var first = View.List.Items.OfType<TextItem>().FirstOrDefault(t => t.ShapeId == id);
        Editor.FontSize = first != null ? Math.Max(10, first.Size * (r.Width / hit.Box.W)) : 14;
        Editor.FontFamily = first != null ? new FontFamily(first.Font) : new FontFamily("Segoe UI");
        Editor.Text = _doc.GetShapeText(_cur, id);
        Editor.Visibility = Visibility.Visible; Editor.Focus(); Editor.SelectAll();
    }

    void CommitEdit()
    {
        if (_editId == 0 || _doc == null) return;
        int id = _editId; _editId = 0;
        var text = Editor.Text; Editor.Visibility = Visibility.Collapsed;
        if (_cancelEdit || text == _doc.GetShapeText(_cur, id)) return;
        _doc.SetShapeText(_cur, id, text); RefreshCurrent();
    }

    void OnEditorLost(object s, RoutedEventArgs e) => CommitEdit();
    void OnEditorKey(object s, KeyEventArgs e)
    {
        if (e.Key == Key.Escape) { _cancelEdit = true; CommitEdit(); View.Focus(); e.Handled = true; }
        else if (e.Key == Key.Enter && Keyboard.Modifiers == ModifierKeys.Control) { CommitEdit(); View.Focus(); e.Handled = true; }
    }

    void OnBoxChanged(int id, RectD b)
    {
        if (_doc == null) return;
        try { _doc.SetShapeBox(_cur, id, b.X, b.Y, b.W, b.H); RefreshCurrent(); }
        catch (Exception ex) when (ex is KeyNotFoundException or NotSupportedException) { Status.Text = ex.Message; }
    }

    TextItem? FirstText(int id) => View.List?.Items.OfType<TextItem>().FirstOrDefault(t => t.ShapeId == id);

    void Format(Func<TextItem, (bool? b, bool? i, bool? u, double? sz)> f, string? color = null)
    {
        int id = View.SelectedId; if (_doc == null || id == 0) { Status.Text = "Seleziona prima una forma con testo."; return; }
        var t = FirstText(id); if (t == null && color == null) { Status.Text = "La forma selezionata non contiene testo."; return; }
        var (b, i, u, sz) = t != null ? f(t) : (null, null, null, null);
        _doc.SetFormat(_cur, id, b, i, u, sz, color); RefreshCurrent();
    }

    bool HasUnderline(int id) => View.List!.Items.OfType<PathItem>().Any(p => p.ShapeId == id && p.Fill == null && p.Segs.Count == 2 && p.Stroke != null && p.Stroke.Width < 3 && FirstText(id) is { } t && p.Stroke.Color == t.Color);

    void OnBold(object s, RoutedEventArgs e) => Format(t => (!t.Bold, null, null, null));
    void OnItalic(object s, RoutedEventArgs e) => Format(t => (null, !t.Italic, null, null));
    void OnUnderline(object s, RoutedEventArgs e) => Format(t => (null, null, !HasUnderline(t.ShapeId), null));
    void OnSizeUp(object s, RoutedEventArgs e) => Format(t => (null, null, null, Math.Round(t.Size) + 2));
    void OnSizeDown(object s, RoutedEventArgs e) => Format(t => (null, null, null, Math.Max(6, Math.Round(t.Size) - 2)));
    void OnColor(object s, RoutedEventArgs e) => Format(_ => (null, null, null, null), (string)((FrameworkElement)s).Tag);

    void OnDeleteShape(object s, RoutedEventArgs e)
    {
        if (_doc == null || View.SelectedId == 0) return;
        _doc.DeleteShape(_cur, View.SelectedId); View.SelectedId = 0; RefreshCurrent();
    }

    void OnInsertText(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        var (w, h) = _doc.SlideSize;
        int id = _doc.AddTextBox(_cur, w / 2 - 150, h / 2 - 20, 300, 40, "Testo");
        RefreshCurrent(); View.SelectedId = id; UpdateLayout(); BeginEdit(id);
    }

    void OnInsertShape(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        int k = ListPicker.Pick(this, "Inserisci forma", Shapes.Select(x => x.Name)); if (k < 0) return;
        var (w, h) = _doc.SlideSize;
        int id = _doc.AddShape(_cur, Shapes[k].Prst, w / 2 - 100, h / 2 - 60, 200, 120);
        RefreshCurrent(); View.SelectedId = id;
    }

    void OnInsertPicture(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        var dlg = new OpenFileDialog { Filter = "Immagini (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp", Title = "Inserisci immagine" };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            var data = File.ReadAllBytes(dlg.FileName);
            var fr = BitmapFrame.Create(new MemoryStream(data), BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
            double pw = fr.PixelWidth * 72.0 / (fr.DpiX > 0 ? fr.DpiX : 96), ph = fr.PixelHeight * 72.0 / (fr.DpiY > 0 ? fr.DpiY : 96);
            var (sw, sh) = _doc.SlideSize; double k = Math.Min(1, Math.Min(sw * 0.6 / pw, sh * 0.6 / ph));
            double w = pw * k, h = ph * k;
            int id = _doc.AddPicture(_cur, data, (sw - w) / 2, (sh - h) / 2, w, h, Path.GetFileNameWithoutExtension(dlg.FileName));
            RefreshCurrent(); View.SelectedId = id;
        }
        catch (Exception ex) when (ex is IOException or NotSupportedException or FileFormatException or PptxFormatException or UnauthorizedAccessException)
        { MessageBox.Show(this, "Immagine non inseribile: " + ex.Message, "Inserisci immagine", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    void OnInsertTable(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        var (w, h) = _doc.SlideSize;
        int id = _doc.AddTable(_cur, 3, 3, w * 0.15, h * 0.35, w * 0.7);
        RefreshCurrent(); View.SelectedId = id;
    }

    // ---------- transizioni ----------
    void UpdateTransUi()
    {
        if (_doc == null) return;
        var t = _doc.GetTransition(_cur);
        foreach (var b in TransPanel.Children.OfType<ToggleButton>()) b.IsChecked = (string)b.Tag == t.Kind;
        TransInfo.Text = t.Kind == "none" ? "" : $"{t.Kind} · direzione {t.Dir ?? "—"} · velocità {t.Speed}";
    }

    void OnTransition(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        var kind = (string)((FrameworkElement)s).Tag; var old = _doc.GetTransition(_cur);
        _doc.SetTransition(_cur, kind, kind is "push" or "wipe" or "cover" or "pull" ? old.Dir ?? "l" : null, old.Speed);
        UpdateTransUi(); UpdateTitle();
    }

    void OnTransDir(object s, RoutedEventArgs e)
    {
        if (_doc == null) return; var t = _doc.GetTransition(_cur);
        if (t.Kind is "push" or "wipe" or "cover" or "pull") { _doc.SetTransition(_cur, t.Kind, (string)((FrameworkElement)s).Tag, t.Speed); UpdateTransUi(); UpdateTitle(); }
    }

    void OnTransSpeed(object s, RoutedEventArgs e)
    {
        if (_doc == null) return; var t = _doc.GetTransition(_cur);
        if (t.Kind != "none") { _doc.SetTransition(_cur, t.Kind, t.Dir, (string)((FrameworkElement)s).Tag); UpdateTransUi(); UpdateTitle(); }
    }

    // ---------- presentazione ----------
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int i);

    void OnPresent(object s, RoutedEventArgs e) => StartShow(0, false);
    void OnPresentHere(object s, RoutedEventArgs e) => StartShow(_cur, false);
    void OnPresenter(object s, RoutedEventArgs e) => StartShow(_cur, true);

    public (PresentationWindow, PresenterWindow?) StartShow(int start, bool withPresenter, bool animate = true)
    {
        CommitEdit(); SaveNotes();
        var c = new ShowController(_doc!, _r!, start);
        var show = new PresentationWindow(c) { Animate = animate }; PresenterWindow? pv = null;
        bool multi = GetSystemMetrics(80) > 1;
        Rect? secondary = null;
        if (multi)
        {
            // ponytail: secondo schermo assunto a destra o a sinistra del principale, stessa altezza (non verificato: PC di sviluppo con un solo schermo)
            double vl = SystemParameters.VirtualScreenLeft, vw = SystemParameters.VirtualScreenWidth, pw = SystemParameters.PrimaryScreenWidth;
            secondary = vl < 0 ? new Rect(vl, 0, -vl, SystemParameters.PrimaryScreenHeight) : new Rect(pw, 0, vw - pw, SystemParameters.PrimaryScreenHeight);
        }
        if (withPresenter)
        {
            pv = new PresenterWindow(c);
            if (secondary == null)
            {
                // un solo schermo: vista relatore a sinistra, presentazione in finestra a destra
                pv.WindowStartupLocation = WindowStartupLocation.Manual; pv.Left = 10; pv.Top = 10; pv.Width = 1000; pv.Height = 700;
                show.WindowStyle = WindowStyle.SingleBorderWindow; show.ResizeMode = ResizeMode.CanResize; show.Cursor = Cursors.Arrow;
                show.WindowStartupLocation = WindowStartupLocation.Manual; show.Left = 1030; show.Top = 60; show.Width = 800; show.Height = 480;
            }
            else pv.WindowState = WindowState.Maximized;
            pv.Show();
        }
        if (secondary != null || !withPresenter) show.ShowFullScreen(secondary); else show.Show();
        return (show, pv);
    }

    // ---------- tema e scorciatoie ----------
    void OnTheme(object s, RoutedEventArgs e)
    {
        _dark = !_dark;
        var r = Application.Current.Resources;
        void Set(string k, string hex) => r[k] = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        if (_dark) { Set("BgBrush", "#0e0e0e"); Set("PanelBrush", "#161616"); Set("FgBrush", "#f2ece1"); Set("MutedBrush", "#9a948a"); Set("LineBrush", "#2e2e2e"); Set("CanvasBrush", "#050505"); Set("HoverBrush", "#222222"); }
        else { Set("BgBrush", "#f4f1ec"); Set("PanelBrush", "#ffffff"); Set("FgBrush", "#1a1a1a"); Set("MutedBrush", "#6b665e"); Set("LineBrush", "#d6d0c6"); Set("CanvasBrush", "#dcd8d0"); Set("HoverBrush", "#ece7de"); }
        View.InvalidateVisual();
    }

    void OnWindowKey(object s, KeyEventArgs e)
    {
        var ctrl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
        if (e.Key == Key.F5 && _doc != null) { StartShow(Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) ? _cur : 0, false); e.Handled = true; }
        else if (ctrl && e.Key == Key.S) { DoSave(Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)); e.Handled = true; }
        else if (ctrl && e.Key == Key.O) { OnOpen(this, e); e.Handled = true; }
        else if (ctrl && e.Key == Key.N) { OnNew(this, e); e.Handled = true; }
        else if (e.Key == Key.Delete && View.IsKeyboardFocused && View.SelectedId != 0) { OnDeleteShape(this, e); e.Handled = true; }
        else if (View.IsKeyboardFocused && e.Key == Key.PageDown && _doc != null) { ShowSlide(_cur + 1); e.Handled = true; }
        else if (View.IsKeyboardFocused && e.Key == Key.PageUp && _doc != null) { ShowSlide(_cur - 1); e.Handled = true; }
    }

    // ---------- schermate per la documentazione ----------
    static void Capture(FrameworkElement el, string path)
    {
        el.UpdateLayout();
        var rtb = new RenderTargetBitmap((int)el.ActualWidth, (int)el.ActualHeight, 96, 96, PixelFormats.Pbgra32);
        rtb.Render(el); SlideDrawing.SavePng(rtb, path);
    }

    public void ScheduleScreenshot(string path, string? mode)
    {
        var t = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
        t.Tick += (_, _) =>
        {
            t.Stop();
            bool wait = false;
            try
            {
                if (mode == "present")
                {
                    var (show, pv) = StartShow(0, true, false);
                    var t2 = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
                    t2.Tick += (_, _) =>
                    {
                        t2.Stop();
                        try { Capture((FrameworkElement)show.Content, Path.ChangeExtension(path, null) + "-presentazione.png"); if (pv != null) Capture((FrameworkElement)pv.Content, Path.ChangeExtension(path, null) + "-relatore.png"); }
                        finally { Application.Current.Shutdown(); }
                    };
                    t2.Start(); wait = true; return;
                }
                if (mode == "edit") { var h = View.List!.Hits.FirstOrDefault(x => x.HasText); if (h.Id != 0) { View.SelectedId = h.Id; BeginEdit(h.Id); UpdateLayout(); } }
                if (mode == "light") OnTheme(this, new RoutedEventArgs());
                if (mode == "sorter") SetSorter(true);
                if (mode?.StartsWith("slide") == true && int.TryParse(mode[5..], out var n)) ShowSlide(n - 1);
                Capture((FrameworkElement)Content, path);
            }
            finally { if (!wait) Application.Current.Shutdown(); }
        };
        t.Start();
    }
}
