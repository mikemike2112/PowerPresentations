using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;
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
    /// <summary>Nome accessibile: numero e primo testo della slide.</summary>
    public string Title { get; set; } = "";
    public override string ToString() => Title;
    public double Opacity { get; set; } = 1;
    public double Width { get; set; } = 170;
    public ImageSource? Image { get => _img; set { _img = value; PropertyChanged?.Invoke(this, new(nameof(Image))); } }
    public event PropertyChangedEventHandler? PropertyChanged;
}

public sealed record SearchHit(string Text, Action Run) { public override string ToString() => Text; }

public partial class MainWindow : Window
{
    PptxDocument? _doc;
    SlideRenderer? _r;
    int _cur;
    bool _loading;
    int _editId; (int Frame, int Row, int Col)? _editCell;
    bool _cancelEdit;
    bool _sorter;
    bool _showNotes = true;
    AppSettings _settings = AppSettings.Load();
    readonly ObservableCollection<ThumbItem> _thumbs = new();
    readonly List<byte[]> _undo = new(), _redo = new();
    Point _dragStart; int _dragIndex = -1;
    TabItem? _lastTab; TextBlock _backInfo = null!;
    Button _undoBtn = null!, _redoBtn = null!;
    Slider _zoom = null!; TextBlock _zoomText = null!;

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
        ApplyTheme(_settings.Dark);
        Thumbs.ItemsSource = _thumbs;
        BuildChrome(); BuildRibbon(); BuildStatus(); BuildBackstage();
        View.Selected += OnSelect;
        View.BoxChanged += OnBoxChanged;
        View.EditRequested += BeginEdit;
        PreviewKeyDown += OnWindowKey;
        Closing += OnClosing;
        StateChanged += (_, _) => Root.Margin = WindowState == WindowState.Maximized ? new Thickness(7) : new Thickness(0);
        Thumbs.MouseDoubleClick += OnThumbDoubleClick;
        Drop += (_, e) => { if (e.Data.GetData(DataFormats.FileDrop) is string[] f && f.Length > 0 && f[0].EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) && Confirm()) OpenFile(f[0]); };
        NewDoc();
    }

    // ---------- tema ----------
    public void ApplyTheme(bool dark)
    {
        var r = Application.Current.Resources;
        void Set(string k, string hex) => r[k] = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        if (dark) { Set("BgBrush", "#0e0e0e"); Set("PanelBrush", "#161616"); Set("FgBrush", "#f2ece1"); Set("MutedBrush", "#9a948a"); Set("LineBrush", "#2e2e2e"); Set("CanvasBrush", "#0e0e0e"); Set("HoverBrush", "#222222"); Set("RibbonBrush", "#161616"); }
        else { Set("BgBrush", "#eee8e6"); Set("PanelBrush", "#ffffff"); Set("FgBrush", "#1f1f1f"); Set("MutedBrush", "#616161"); Set("LineBrush", "#d6cfcd"); Set("CanvasBrush", "#eee8e6"); Set("HoverBrush", "#e4dcda"); Set("RibbonBrush", "#fbf9f8"); }
        View?.InvalidateVisual();
    }

    void SetDark(bool dark)
    {
        _settings.Dark = dark; _settings.Save(); ApplyTheme(dark);
    }

    // ---------- barra del titolo, ribbon, stato ----------
    void BuildChrome()
    {
        SearchIcon.Content = Icons.Make("search", 16);
        var ico = new Image { Source = Icon, Width = 20, Height = 20, Margin = new Thickness(0, 0, 8, 0) };
        QuickBar.Children.Add(ico);
        Button Q(string name, string icon, Action run) { var b = new Button { Style = (Style)FindResource("RibbonButton"), Width = 32, Height = 30, Padding = new Thickness(0), Content = Icons.Make(icon, 18), ToolTip = name }; AutomationProperties.SetName(b, name); b.Click += (_, _) => run(); QuickBar.Children.Add(b); return b; }
        Q("Salva", "save", () => DoSave(false));
        _undoBtn = Q("Annulla", "undo", DoUndo); _redoBtn = Q("Ripristina azione", "redo", DoRedo);
        Q("Avvia presentazione", "play-here", () => StartShow(0, false));
        UpdateQuick();
        Button C(string name, string icon, Action run, bool close = false)
        {
            var b = new Button { Style = (Style)FindResource("RibbonButton"), Width = 46, Height = 40, Padding = new Thickness(0), Content = Icons.Make(icon, 13), ToolTip = name };
            AutomationProperties.SetName(b, name); b.Click += (_, _) => run(); CaptionButtons.Children.Add(b); return b;
        }
        C("Riduci a icona", "min", () => WindowState = WindowState.Minimized);
        _maxBtn = C("Ingrandisci", "max", () => WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized);
        StateChanged += (_, _) => UpdateMaxButton();   // anche con doppio clic sulla barra e Win+frecce
        C("Chiudi", "close", Close);

        var collapse = new Button { Style = (Style)FindResource("RibbonButton"), Width = 26, Height = 22, Padding = new Thickness(0), Content = Icons.Make("dir-u", 12), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(0, 8, 14, 0), ToolTip = "Comprimi il ribbon" };
        AutomationProperties.SetName(collapse, "Comprimi il ribbon");
        collapse.Click += (_, _) =>
        {
            Ribbon.ApplyTemplate();
            if (Ribbon.Template.FindName("RibbonBody", Ribbon) is Border b) { bool vis = b.Visibility == Visibility.Visible; b.Visibility = vis ? Visibility.Collapsed : Visibility.Visible; collapse.Content = Icons.Make(vis ? "dir-d" : "dir-u", 12); }
        };
        Grid.SetRow(collapse, 1); Root.Children.Add(collapse);
    }

    Button _maxBtn = null!;

    /// <summary>Un quadrato quando si può ingrandire, due quadrati sovrapposti quando si può ripristinare.</summary>
    void UpdateMaxButton()
    {
        bool max = WindowState == WindowState.Maximized;
        _maxBtn.Content = Icons.Make(max ? "restore" : "max", 13);
        string t = max ? "Ripristina giù" : "Ingrandisci"; _maxBtn.ToolTip = t; AutomationProperties.SetName(_maxBtn, t);
    }

    /// <summary>Compatta titolo, schede e barra di stato quando la finestra si stringe.</summary>
    void AdaptChrome()
    {
        double w = ActualWidth;
        double pad = w < 760 ? 5 : w < 900 ? 7 : 11; double fs = w < 760 ? 12.5 : w < 900 ? 13 : 14;
        foreach (var t in Ribbon.Items.OfType<TabItem>()) { t.Padding = new Thickness(pad, 6, pad, 6); t.FontSize = fs; }
        if (_zoom != null) _zoom.Width = w < 900 ? 70 : 130;
    }

    void UpdateQuick() { if (_undoBtn != null) { _undoBtn.IsEnabled = _undo.Count > 0; _redoBtn.IsEnabled = _redo.Count > 0; } }

    void BuildStatus()
    {
        Button B(string name, string icon, Action run) { var b = new Button { Style = (Style)FindResource("RibbonButton"), Width = 30, Height = 22, Padding = new Thickness(0), Content = Icons.Make(icon, 14), ToolTip = name }; AutomationProperties.SetName(b, name); b.Click += (_, _) => run(); StatusRight.Children.Add(b); return b; }
        var notes = new ToggleButton { Content = "Note", IsChecked = true, FontSize = 12, Padding = new Thickness(8, 1, 8, 1) }; AutomationProperties.SetName(notes, "Note (barra di stato)");
        notes.Click += (_, _) => SetNotesVisible(notes.IsChecked == true); StatusRight.Children.Add(notes);
        B("Normale", "normal", () => SetSorter(false)); B("Sequenza diapositive", "sorter", () => SetSorter(true)); B("Presentazione", "play-here", () => StartShow(_cur, false));
        B("Riduci zoom", "min", () => SetZoom((int)_zoom.Value - 10));
        _zoom = new Slider { Minimum = 25, Maximum = 200, Value = 100, Width = 130, VerticalAlignment = VerticalAlignment.Center }; AutomationProperties.SetName(_zoom, "Zoom");
        _zoom.ValueChanged += (_, e) => { View.Zoom = e.NewValue / 100; _zoomText.Text = $"{(int)e.NewValue}%"; };
        StatusRight.Children.Add(_zoom);
        B("Aumenta zoom", "slide-new", () => SetZoom((int)_zoom.Value + 10));
        _zoomText = new TextBlock { Text = "100%", Width = 42, TextAlignment = TextAlignment.Right, VerticalAlignment = VerticalAlignment.Center, FontSize = 12, Margin = new Thickness(4, 0, 4, 0) }; StatusRight.Children.Add(_zoomText);
        B("Adatta alla finestra", "fit", () => SetZoom(100));
    }

    void SetZoom(int v) => _zoom.Value = Math.Clamp(v, 25, 200);

    void SetNotesVisible(bool on)
    {
        _showNotes = on; NotesRow.Height = on ? new GridLength(110) : new GridLength(0); NotesArea.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
        if (_notesCheck != null) { _loading = true; _notesCheck.IsChecked = on; _loading = false; }
        foreach (var t in StatusRight.Children.OfType<ToggleButton>()) t.IsChecked = on;
    }

    void OnRibbonTab(object s, SelectionChangedEventArgs e)
    {
        if (e.OriginalSource != Ribbon || Ribbon.SelectedItem is not TabItem t) return;
        if ((t.Header as string) == "File") { ShowBackstage(true); return; }
        ShowBackstage(false); _lastTab = t;
        Ribbon.ApplyTemplate();
        if (Ribbon.Template.FindName("RibbonBody", Ribbon) is Border b)
        {
            bool powers = (t.Header as string) == "Powers";
            b.SetResourceReference(Border.BorderBrushProperty, powers ? "PowersBrush" : "LineBrush"); b.BorderThickness = new Thickness(powers ? 2 : 1);
        }
    }

    // ---------- scheda File a tutta pagina ----------
    void BuildBackstage()
    {
        Backstage.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(260) }); Backstage.ColumnDefinitions.Add(new ColumnDefinition());
        var left = new StackPanel { Margin = new Thickness(12, 12, 8, 0) };
        Button M(string label, string icon, Action? run)
        {
            var b = new Button { Style = (Style)FindResource("RibbonButton"), Height = 40, HorizontalContentAlignment = HorizontalAlignment.Left, Padding = new Thickness(12, 0, 8, 0) };
            var sp = new StackPanel { Orientation = Orientation.Horizontal }; sp.Children.Add(Icons.Make(icon, 20)); sp.Children.Add(new TextBlock { Text = label, FontSize = 14, Margin = new Thickness(12, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center });
            b.Content = sp; AutomationProperties.SetName(b, label);
            if (run == null) { b.IsEnabled = false; b.ToolTip = NotYet; } else { b.Click += (_, _) => { ShowBackstage(false); Ribbon.SelectedItem = _lastTab ?? _homeTab; run(); }; }
            left.Children.Add(b); return b;
        }
        var back = M("Indietro", "back", null); back.IsEnabled = true; back.ToolTip = null;
        back.Click += (_, _) => { ShowBackstage(false); Ribbon.SelectedItem = _lastTab ?? _homeTab; };
        M("Nuovo", "new", () => OnNew(this, new RoutedEventArgs())); M("Apri", "open", () => OnOpen(this, new RoutedEventArgs())); M("Salva", "save", () => DoSave(false)); M("Salva con nome", "save", () => DoSave(true));
        M("Esporta come PDF", "pdf", () => OnExportPdf(this, new RoutedEventArgs())); M("Stampa", "generic", null);
        Backstage.Children.Add(left);
        var info = new StackPanel { Margin = new Thickness(40, 40, 0, 0) }; Grid.SetColumn(info, 1);
        info.Children.Add(new TextBlock { Text = "Informazioni", FontSize = 28 });
        _backInfo = new TextBlock { FontSize = 14, Margin = new Thickness(0, 14, 0, 0) }; info.Children.Add(_backInfo);
        Backstage.Children.Add(info);
    }

    void ShowBackstage(bool on)
    {
        Backstage.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
        if (!on) Dispatcher.BeginInvoke(() => View.Focus(), DispatcherPriority.Input);   // la tastiera (Ctrl+Z...) torna all'area di lavoro
        if (on && Backstage.Children.OfType<StackPanel>().Last().Children.OfType<TextBlock>().Last() is { } t)
            t.Text = $"{(_doc?.FilePath ?? "Presentazione non ancora salvata")}\nDiapositive: {_doc?.Slides.Count}";
    }

    // ---------- ricerca ----------
    void OnSearchFocus(object s, RoutedEventArgs e) { if (Search.Text.Length > 0) SearchPopup.IsOpen = true; }

    void OnSearchChanged(object s, TextChangedEventArgs e)
    {
        SearchHint.Visibility = Search.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
        var q = Search.Text.Trim();
        if (q.Length == 0 || _doc == null) { SearchPopup.IsOpen = false; return; }
        var hits = new List<SearchHit>();
        foreach (var (label, run) in _cmds.DistinctBy(c => c.Label).Where(c => c.Label.Contains(q, StringComparison.CurrentCultureIgnoreCase)).Take(8)) hits.Add(new SearchHit("Comando: " + label, run));
        if (q.Length >= 2)
            for (int i = 0; i < _doc.Slides.Count && hits.Count < 14; i++)
            {
                var txt = string.Join(" ", _r!.Render(i).Items.OfType<TextItem>().Select(t => t.Text));
                int idx = txt.IndexOf(q, StringComparison.CurrentCultureIgnoreCase); if (idx < 0) continue;
                int n = i; var snip = txt.Substring(Math.Max(0, idx - 15), Math.Min(txt.Length - Math.Max(0, idx - 15), 50));
                hits.Add(new SearchHit($"Diapositiva {i + 1}: …{snip}…", () => { SetSorter(false); ShowSlide(n); }));
            }
        if (hits.Count == 0) hits.Add(new SearchHit("Nessun risultato", () => { }));
        SearchList.ItemsSource = hits; SearchPopup.IsOpen = true;
    }

    void RunHit(SearchHit? h) { if (h == null) return; SearchPopup.IsOpen = false; Search.Text = ""; h.Run(); }
    void OnSearchPick(object s, MouseButtonEventArgs e) { if (SearchList.SelectedItem is SearchHit h) RunHit(h); }
    void OnSearchKey(object s, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) { RunHit((SearchList.SelectedItem as SearchHit) ?? (SearchList.Items.Count > 0 ? SearchList.Items[0] as SearchHit : null)); e.Handled = true; }
        else if (e.Key == Key.Down && SearchList.Items.Count > 0) { SearchList.SelectedIndex = Math.Min(SearchList.SelectedIndex + 1, SearchList.Items.Count - 1); e.Handled = true; }
        else if (e.Key == Key.Up && SearchList.Items.Count > 0) { SearchList.SelectedIndex = Math.Max(SearchList.SelectedIndex - 1, 0); e.Handled = true; }
        else if (e.Key == Key.Escape) { SearchPopup.IsOpen = false; Search.Text = ""; View.Focus(); e.Handled = true; }
    }

    // ---------- annulla / ripristina ----------
    void Snap()
    {
        if (_doc == null) return;
        _undo.Add(_doc.Snapshot()); if (_undo.Count > 40) _undo.RemoveAt(0);
        _redo.Clear(); UpdateQuick();
    }

    void ApplyState(byte[] state)
    {
        _doc!.Restore(state); _r = new SlideRenderer(_doc); _cur = Math.Min(_cur, _doc.Slides.Count - 1);
        _editId = 0; _editCell = null; Editor.Visibility = Visibility.Collapsed;
        RebuildThumbs(); ShowSlide(_cur); UpdateTitle(); UpdateQuick();
        View.Focus();   // la ricostruzione delle miniature perde il fuoco: senza, Ctrl+Y dopo Ctrl+Z non arrivava
    }

    void DoUndo()
    {
        if (_doc == null || _undo.Count == 0) return;
        CommitEdit(); var s = _undo[^1]; _undo.RemoveAt(_undo.Count - 1); _redo.Add(_doc.Snapshot()); ApplyState(s);
    }

    void DoRedo()
    {
        if (_doc == null || _redo.Count == 0) return;
        CommitEdit(); var s = _redo[^1]; _redo.RemoveAt(_redo.Count - 1); _undo.Add(_doc.Snapshot()); ApplyState(s);
    }

    // ---------- documento ----------
    void NewDoc()
    {
        try { SetDoc(PptxDocument.New()); }
        catch (Exception e) when (e is InvalidOperationException or PptxFormatException) { ThemedDialog.Notify(this, "Nuova presentazione", e.Message, DialogKind.Warning); }
    }

    public void OpenFile(string path)
    {
        try { SetDoc(PptxDocument.Open(path)); }
        catch (PptxFormatException e) { ThemedDialog.Notify(this, "Impossibile aprire il file", e.Message, DialogKind.Warning); }
    }

    void SetDoc(PptxDocument d)
    {
        _doc?.Dispose(); _doc = d; _r = new SlideRenderer(d); _cur = 0; _editId = 0; _editCell = null; Editor.Visibility = Visibility.Collapsed;
        _undo.Clear(); _redo.Clear(); UpdateQuick();
        RebuildThumbs(); ShowSlide(0); UpdateTitle();
    }

    /// <summary>Richiesta di salvataggio prima di chiudere, Nuovo o Apri. Ritorna false se l'utente annulla (o il salvataggio fallisce).</summary>
    bool Confirm()
    {
        if (_doc is not { Dirty: true }) return true;
        var (choice, path) = ThemedDialog.AskSave(this, _doc.FilePath, false);
        if (!SavePlan.Proceeds(choice)) return false;
        return !SavePlan.Writes(choice) || SaveTo(path!);
    }

    void OnClosing(object? s, CancelEventArgs e) { CommitEdit(); SaveNotes(); if (!Confirm()) e.Cancel = true; else _doc?.Dispose(); }

    void UpdateTitle()
    {
        var t = $"{(_doc?.FilePath is { } p ? Path.GetFileName(p) : "Presentazione1")}{(_doc?.Dirty == true ? " *" : "")} — PowerPresentations";
        Title = t; TitleText.Text = t;
    }

    void UpdateStatus() =>
        Status.Text = _doc == null ? "" : $"Diapositiva {_cur + 1} di {_doc.Slides.Count}" + (View.SelectedId != 0 ? $"   ·   forma {View.SelectedId}" : "") + (_doc.IsHidden(_cur) ? "   ·   nascosta" : "");

    // ---------- selezione / scheda tabella ----------
    bool IsTable(int id) => id != 0 && View.List != null && View.List.TableCells.Any(c => c.FrameId == id);

    void OnSelect(int id)
    {
        UpdateStatus();
        bool table = IsTable(id);
        _tableTab.Visibility = table ? Visibility.Visible : Visibility.Collapsed;
        if (!table && Ribbon.SelectedItem == _tableTab) Ribbon.SelectedItem = _homeTab;
        if (table && _doc != null)
        {
            _loading = true;
            foreach (var c in _tableChecks) c.IsChecked = _doc.GetTableFlag(_cur, id, FlagOf(c));
            _loading = false;
        }
    }

    static string FlagOf(CheckBox c) => AutomationProperties.GetName(c) switch
    {
        "Riga di intestazione" => "firstRow", "Riga Totale" => "lastRow", "Righe alternate evidenziate" => "bandRow",
        "Prima colonna" => "firstCol", "Ultima colonna" => "lastCol", _ => "bandCol"
    };

    void SetTableFlag(string flag, bool on)
    {
        int id = View.SelectedId; if (_doc == null || !IsTable(id)) return;
        Snap(); _doc.SetTableFlag(_cur, id, flag, on); RefreshCurrent(); View.SelectedId = id;
    }

    // ---------- miniature ----------
    ThumbItem MakeThumb(int i)
    {
        var bmp = SlideDrawing.ToBitmap(_r!.Render(i), 320);
        var first = _r.Render(i).Items.OfType<TextItem>().FirstOrDefault()?.Text ?? "";
        return new ThumbItem { Title = $"Diapositiva {i + 1}: {first}", Number = i + 1, Image = bmp, Opacity = _doc!.IsHidden(i) ? 0.4 : 1, Width = _sorter ? 300 : 170 };
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
        View.SelectedId = 0; View.List = _r!.Render(_cur); OnSelect(0);
        Notes.Text = _doc.GetNotes(_cur);
        UpdateTransUi(); UpdateStatus();
        if (Thumbs.SelectedIndex != _cur) { _loading = true; Thumbs.SelectedIndex = _cur; _loading = false; }
    }

    void OnThumbSelected(object s, SelectionChangedEventArgs e) { if (!_loading && Thumbs.SelectedIndex >= 0) { SaveNotes(); ShowSlide(Thumbs.SelectedIndex); } }

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
        SaveNotes(); Snap();
        _doc.MoveSlide(from, to); _cur = to; RebuildThumbs(); ShowSlide(to); UpdateTitle();
    }

    void OnThumbDoubleClick(object s, MouseButtonEventArgs e) { if (_sorter && Thumbs.SelectedIndex >= 0) { int i = Thumbs.SelectedIndex; SetSorter(false); ShowSlide(i); } }

    void OnThumbKey(object s, KeyEventArgs e) { if (e.Key == Key.Delete) { OnDeleteSlide(s, e); e.Handled = true; } }

    void SetSorter(bool on)
    {
        if (_doc == null) return;
        CommitEdit(); SaveNotes(); _sorter = on;
        Grid.SetColumnSpan(Thumbs, on ? 2 : 1); Grid.SetRowSpan(Thumbs, 2);
        EditArea.Visibility = on ? Visibility.Collapsed : Visibility.Visible;
        NotesArea.Visibility = !on && _showNotes ? Visibility.Visible : Visibility.Collapsed;
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

    bool DoSave(bool saveAs)
    {
        if (_doc == null) return false;
        CommitEdit(); SaveNotes();
        string? path = _doc.FilePath;
        if (saveAs || path == null)
        {
            var (choice, chosen) = ThemedDialog.AskSave(this, path, true);
            if (!SavePlan.Writes(choice)) return false;
            path = chosen;
        }
        return SaveTo(path!);
    }

    bool SaveTo(string path)
    {
        try { _doc!.SaveAs(path); UpdateTitle(); return true; }   // atomico: file temporaneo + sostituzione
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
        {
            ThemedDialog.Notify(this, "Salvataggio non riuscito", "Il file originale non è stato modificato.\n" + ex.Message, DialogKind.Error);
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
        { ThemedDialog.Notify(this, "Esportazione PDF non riuscita", ex.Message, DialogKind.Error); }
    }

    // ---------- slide ----------
    void OnNewSlide(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        var lays = _doc.Layouts;
        int k = ListPicker.Pick(this, "Layout della nuova slide", lays.Select(l => l.Name));
        if (k < 0) return;
        SaveNotes(); Snap();
        _cur = _doc.AddSlide(lays[k], _cur); RebuildThumbs(); ShowSlide(_cur); UpdateTitle();
    }

    void OnDuplicate(object s, RoutedEventArgs e) { if (_doc == null) return; SaveNotes(); Snap(); _cur = _doc.DuplicateSlide(_cur); RebuildThumbs(); ShowSlide(_cur); UpdateTitle(); }

    void OnDeleteSlide(object s, RoutedEventArgs e)
    {
        if (_doc == null || _doc.Slides.Count <= 1) { Status.Text = "Non si può eliminare l'ultima slide."; return; }
        Snap(); _doc.DeleteSlide(_cur); _cur = Math.Min(_cur, _doc.Slides.Count - 1); RebuildThumbs(); ShowSlide(_cur); UpdateTitle();
    }

    void OnToggleHidden(object s, RoutedEventArgs e) { if (_doc == null) return; Snap(); _doc.SetHidden(_cur, !_doc.IsHidden(_cur)); RefreshCurrent(); UpdateStatus(); }

    void SaveNotes()
    {
        if (_doc == null || _cur >= _doc.Slides.Count) return;
        if (Notes.Text != _doc.GetNotes(_cur)) { Snap(); _doc.SetNotes(_cur, Notes.Text); UpdateTitle(); }
    }

    void OnNotesLost(object s, RoutedEventArgs e) => SaveNotes();
    void OnNotesChanged(object s, TextChangedEventArgs e) => NotesHint.Visibility = Notes.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;

    // ---------- modifica testo e oggetti ----------
    void BeginEdit(int id, Point slidePt)
    {
        if (_doc == null) return;
        var hit = View.List!.Hits.FirstOrDefault(h => h.Id == id); if (hit.Id == 0) return;
        _editId = id; _cancelEdit = false; _editCell = null;
        RectD box = hit.Box; string text;
        TextItem? first;
        var cell = View.List.TableCells.Where(c => c.FrameId == id && slidePt.X >= c.Box.X && slidePt.X <= c.Box.X + c.Box.W && slidePt.Y >= c.Box.Y && slidePt.Y <= c.Box.Y + c.Box.H).Select(c => ((int, int, int, RectD)?)(c.FrameId, c.Row, c.Col, c.Box)).FirstOrDefault();
        if (IsTable(id))
        {
            if (cell == null) { _editId = 0; return; }
            _editCell = (id, cell.Value.Item2, cell.Value.Item3); box = cell.Value.Item4; text = _doc.GetCellText(_cur, id, cell.Value.Item2, cell.Value.Item3);
            first = View.List.Items.OfType<TextItem>().FirstOrDefault(t => t.ShapeId == id);
        }
        else { text = _doc.GetShapeText(_cur, id); first = View.List.Items.OfType<TextItem>().FirstOrDefault(t => t.ShapeId == id); }
        var r = View.ToView(box);
        Editor.Margin = new Thickness(r.X, r.Y, 0, 0); Editor.Width = Math.Max(80, r.Width); Editor.MinHeight = Math.Max(28, r.Height);
        Editor.FontSize = first != null ? Math.Max(10, first.Size * (r.Width / box.W)) : 14;
        Editor.FontFamily = first != null ? new FontFamily(first.Font) : new FontFamily("Segoe UI");
        Editor.Text = text;
        Editor.Visibility = Visibility.Visible; Editor.Focus(); Editor.SelectAll();
    }

    void BeginEdit(int id) => BeginEdit(id, new Point(-1, -1));

    void CommitEdit()
    {
        if (_editId == 0 || _doc == null) return;
        int id = _editId; var cell = _editCell; _editId = 0; _editCell = null;
        var text = Editor.Text; Editor.Visibility = Visibility.Collapsed;
        if (_cancelEdit) return;
        if (cell is { } c)
        {
            if (text == _doc.GetCellText(_cur, c.Frame, c.Row, c.Col)) return;
            Snap(); _doc.SetCellText(_cur, c.Frame, c.Row, c.Col, text);
        }
        else
        {
            if (text == _doc.GetShapeText(_cur, id)) return;
            Snap(); _doc.SetShapeText(_cur, id, text);
        }
        RefreshCurrent(); View.SelectedId = id;
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
        try { Snap(); _doc.SetShapeBox(_cur, id, b.X, b.Y, b.W, b.H); RefreshCurrent(); View.SelectedId = id; }
        catch (Exception ex) when (ex is KeyNotFoundException or NotSupportedException) { if (_undo.Count > 0) _undo.RemoveAt(_undo.Count - 1); Status.Text = ex.Message; }
    }

    TextItem? FirstText(int id) => View.List?.Items.OfType<TextItem>().FirstOrDefault(t => t.ShapeId == id);

    void Format(Func<TextItem, (bool? b, bool? i, bool? u, double? sz)> f, string? color = null)
    {
        int id = View.SelectedId; if (_doc == null || id == 0) { Status.Text = "Seleziona prima una forma con testo."; return; }
        var t = FirstText(id); if (t == null && color == null) { Status.Text = "La forma selezionata non contiene testo."; return; }
        var (b, i, u, sz) = t != null ? f(t) : (null, null, null, null);
        Snap(); _doc.SetFormat(_cur, id, b, i, u, sz, color); RefreshCurrent(); View.SelectedId = id;
    }

    bool HasUnderline(int id) => View.List!.Items.OfType<PathItem>().Any(p => p.ShapeId == id && p.Fill == null && p.Segs.Count == 2 && p.Stroke != null && p.Stroke.Width < 3 && FirstText(id) is { } t && p.Stroke.Color == t.Color);

    void OnBold(object s, RoutedEventArgs e) => Format(t => (!t.Bold, null, null, null));
    void OnItalic(object s, RoutedEventArgs e) => Format(t => (null, !t.Italic, null, null));
    void OnUnderline(object s, RoutedEventArgs e) => Format(t => (null, null, !HasUnderline(t.ShapeId), null));
    void OnSizeUp(object s, RoutedEventArgs e) => Format(t => (null, null, null, Math.Round(t.Size) + 2));
    void OnSizeDown(object s, RoutedEventArgs e) => Format(t => (null, null, null, Math.Max(6, Math.Round(t.Size) - 2)));
    void FormatColor(string hex) => Format(_ => (null, null, null, null), hex);

    void Align(string algn)
    {
        int id = View.SelectedId; if (_doc == null || id == 0 || FirstText(id) == null) { Status.Text = "Seleziona prima una forma con testo."; return; }
        Snap(); _doc.SetAlignment(_cur, id, algn); RefreshCurrent(); View.SelectedId = id;
    }

    void OnDeleteShape(object s, RoutedEventArgs e)
    {
        if (_doc == null || View.SelectedId == 0) return;
        Snap(); _doc.DeleteShape(_cur, View.SelectedId); View.SelectedId = 0; RefreshCurrent(); OnSelect(0);
    }

    void OnInsertText(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        var (w, h) = _doc.SlideSize; Snap();
        int id = _doc.AddTextBox(_cur, w / 2 - 150, h / 2 - 20, 300, 40, "Testo");
        RefreshCurrent(); View.SelectedId = id; UpdateLayout(); BeginEdit(id);
    }

    void OnInsertShape(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        int k = ListPicker.Pick(this, "Inserisci forma", Shapes.Select(x => x.Name)); if (k < 0) return;
        var (w, h) = _doc.SlideSize; Snap();
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
            Snap();
            int id = _doc.AddPicture(_cur, data, (sw - w) / 2, (sh - h) / 2, w, h, Path.GetFileNameWithoutExtension(dlg.FileName));
            RefreshCurrent(); View.SelectedId = id;
        }
        catch (Exception ex) when (ex is IOException or NotSupportedException or FileFormatException or PptxFormatException or UnauthorizedAccessException)
        { ThemedDialog.Notify(this, "Immagine non inseribile", ex.Message, DialogKind.Warning); }
    }

    void OnInsertTable(object s, RoutedEventArgs e)
    {
        if (_doc == null) return;
        var (w, h) = _doc.SlideSize; Snap();
        int id = _doc.AddTable(_cur, 3, 3, w * 0.15, h * 0.35, w * 0.7);
        RefreshCurrent(); View.SelectedId = id; OnSelect(id);
    }

    // ---------- transizioni ----------
    void UpdateTransUi()
    {
        if (_doc == null) return;
        var t = _doc.GetTransition(_cur);
        foreach (var b in _transButtons) b.IsChecked = AutomationProperties.GetName(b) == KindLabel(t.Kind);
    }

    static string KindLabel(string kind) => kind switch { "fade" => "Dissolvenza", "push" => "Spingi", "wipe" => "Cancella", "cover" => "Copri", "pull" => "Scopri", _ => "Nessuna" };

    void SetTrans(string kind)
    {
        if (_doc == null) return;
        var old = _doc.GetTransition(_cur); Snap();
        _doc.SetTransition(_cur, kind, kind is "push" or "wipe" or "cover" or "pull" ? old.Dir ?? "l" : null, old.Speed);
        UpdateTransUi(); UpdateTitle();
    }

    void SetTransDir(string dir)
    {
        if (_doc == null) return; var t = _doc.GetTransition(_cur);
        if (t.Kind is "push" or "wipe" or "cover" or "pull") { Snap(); _doc.SetTransition(_cur, t.Kind, dir, t.Speed); UpdateTransUi(); UpdateTitle(); }
        else Status.Text = "Scegli prima una transizione con direzione (Spingi, Cancella, Copri, Scopri).";
    }

    void SetTransSpeed(string speed)
    {
        if (_doc == null) return; var t = _doc.GetTransition(_cur);
        if (t.Kind != "none") { Snap(); _doc.SetTransition(_cur, t.Kind, t.Dir, speed); UpdateTransUi(); UpdateTitle(); }
        else Status.Text = "Scegli prima una transizione.";
    }

    /// <summary>Anteprima: mostra la slide precedente e passa alla corrente con la sua transizione.</summary>
    void OnTransPreview()
    {
        if (_doc == null || _cur == 0) { Status.Text = "L'anteprima richiede una slide precedente."; return; }
        CommitEdit(); SaveNotes();
        var c = new ShowController(_doc, _r!, _cur - 1);
        var show = new PresentationWindow(c) { WindowStyle = WindowStyle.SingleBorderWindow, ResizeMode = ResizeMode.CanResize, Width = 960, Height = 560, WindowStartupLocation = WindowStartupLocation.CenterOwner, Owner = this, Cursor = Cursors.Arrow };
        show.Show();
        var t = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(700) };
        t.Tick += (_, _) => { t.Stop(); c.Next(); var t2 = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2.2) }; t2.Tick += (_, _) => { t2.Stop(); show.Close(); }; t2.Start(); };
        t.Start();
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

    void OnWindowKey(object s, KeyEventArgs e)
    {
        var ctrl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
        bool typing = Keyboard.FocusedElement is TextBox;
        if (e.Key == Key.F5 && _doc != null) { StartShow(Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) ? _cur : 0, false); e.Handled = true; }
        else if (ctrl && e.Key == Key.S) { DoSave(Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)); e.Handled = true; }
        else if (ctrl && e.Key == Key.O) { OnOpen(this, e); e.Handled = true; }
        else if (ctrl && e.Key == Key.N) { OnNew(this, e); e.Handled = true; }
        else if (ctrl && e.Key == Key.Z && !typing) { DoUndo(); e.Handled = true; }
        else if (ctrl && e.Key == Key.Y && !typing) { DoRedo(); e.Handled = true; }
        else if (ctrl && e.Key == Key.E) { Search.Focus(); e.Handled = true; }
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

    /// <summary>Cattura l'intera finestra (ombra e angoli arrotondati compresi) e ne scrive la misura reale in un file accanto.</summary>
    static void CaptureWindow(Window w, string path)
    {
        w.UpdateLayout();
        var rtb = new RenderTargetBitmap((int)Math.Ceiling(w.ActualWidth), (int)Math.Ceiling(w.ActualHeight), 96, 96, PixelFormats.Pbgra32);
        rtb.Render(w); SlideDrawing.SavePng(rtb, path);
        var body = (FrameworkElement)((Border)w.Content).Child;
        File.WriteAllText(Path.Combine(Path.GetTempPath(), "pp-dialog-misure.txt"), $"finestra {w.ActualWidth}x{w.ActualHeight}; scheda {((FrameworkElement)w.Content).ActualWidth}x{((FrameworkElement)w.Content).ActualHeight}; contenuto {body.ActualWidth}x{body.ActualHeight}; margine interno {body.Margin.Left}/{body.Margin.Right}");
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
                if (mode == "light") { _darkSwitch.IsChecked = false; }
                if (mode == "powers" || mode == "powers-light") { if (mode == "powers-light") _darkSwitch.IsChecked = false; Ribbon.SelectedItem = Ribbon.Items.OfType<TabItem>().First(i => (string)i.Header == "Powers"); }
                if (mode == "backstage") Ribbon.SelectedItem = Ribbon.Items.OfType<TabItem>().First(i => (string)i.Header == "File");
                if (mode == "table")
                {
                    for (int i = 0; i < _doc!.Slides.Count; i++) if (_r!.Render(i).TableCells.Count > 0) { ShowSlide(i); break; }
                    var f = View.List!.TableCells.FirstOrDefault(); if (f.FrameId != 0) { View.SelectedId = f.FrameId; OnSelect(f.FrameId); Ribbon.SelectedItem = _tableTab; }
                }
                if (mode == "privacy")
                {
                    var pd = ShowPrivacyForShot();
                    var t5 = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
                    t5.Tick += (_, _) => { t5.Stop(); try { CaptureWindow(pd, path); } finally { Application.Current.Shutdown(); } };
                    t5.Start(); wait = true; return;
                }
                if (mode == "dialog-live")
                {
                    // finestra viva per la cattura dello schermo vero (tools/dialog-screen.ps1): resta aperta ~12 s poi l'app si chiude
                    NewDoc(); _doc!.MarkDirty();
                    var live = ThemedDialog.BuildSave(this, null, false, "Presentazione", out _); live.Topmost = true; live.Show();
                    var t4 = new DispatcherTimer { Interval = TimeSpan.FromSeconds(12) }; t4.Tick += (_, _) => { t4.Stop(); Application.Current.Shutdown(); }; t4.Start();
                    wait = true; return;
                }
                if (mode is "dialog" or "dialog-light")
                {
                    if (mode == "dialog-light") _darkSwitch.IsChecked = false;
                    NewDoc(); _doc!.MarkDirty();
                    var dlg = ThemedDialog.BuildSave(this, null, false, "Presentazione", out _); dlg.Show();
                    var t3 = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
                    t3.Tick += (_, _) => { t3.Stop(); try { CaptureWindow(dlg, path); } finally { Application.Current.Shutdown(); } };
                    t3.Start(); wait = true; return;
                }
                if (mode == "sorter") SetSorter(true);
                if (mode?.StartsWith("slide") == true && int.TryParse(mode[5..], out var n)) ShowSlide(n - 1);
                UpdateLayout();
                Capture((FrameworkElement)Content, path);
            }
            finally { if (!wait) Application.Current.Shutdown(); }
        };
        t.Start();
    }
}
