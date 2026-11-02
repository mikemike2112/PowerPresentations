using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace PowerPresentations.App;

/// <summary>Costruzione del ribbon (schede, gruppi, comandi grandi/piccoli) con icone proprie. Un comando non implementato è visibile ma disabilitato.</summary>
public partial class MainWindow
{
    readonly List<(string Label, Action Run)> _cmds = new();
    readonly List<ToggleButton> _transButtons = new();
    TabItem _tableTab = null!, _homeTab = null!;
    readonly List<CheckBox> _tableChecks = new();
    ToggleButton _darkSwitch = null!;
    CheckBox _notesCheck = null!;
    const string NotYet = "Non ancora disponibile";

    sealed record CmdMeta(string Label, string Icon, Action? Run);
    sealed class GroupInfo { public FrameworkElement Element = null!; public string Name = ""; public List<CmdMeta> Cmds = new(); public List<TextBlock> Labels = new(); public bool Pinned; }
    readonly List<CmdMeta> _pendingCmds = new(); readonly List<TextBlock> _pendingLabels = new(); bool _pendingPinned;
    readonly Dictionary<FrameworkElement, GroupInfo> _groupInfo = new();
    readonly Dictionary<TabItem, (StackPanel Panel, List<GroupInfo> Groups, Button More)> _tabLayout = new();

    static Brush Res(string k) => (Brush)Application.Current.Resources[k];

    FrameworkElement Wire(ButtonBase b, string label, Action? run, string icon)
    {
        AutomationProperties.SetName(b, label); _pendingCmds.Add(new CmdMeta(label, icon, run));
        if (run == null) { b.IsEnabled = false; b.ToolTip = NotYet; }
        else { b.Click += (_, _) => run(); b.ToolTip = label; _cmds.Add((label, run)); }
        return b;
    }

    FrameworkElement Large(string label, string icon, Action? run)
    {
        var b = new Button { Style = (Style)FindResource("RibbonButton"), Padding = new Thickness(5, 4, 5, 2), MinWidth = 54, VerticalAlignment = VerticalAlignment.Stretch };
        var sp = new StackPanel { Width = double.NaN };
        sp.Children.Add(Icons.Make(icon, 30));
        var lbl = new TextBlock { Text = label, TextWrapping = TextWrapping.Wrap, TextAlignment = TextAlignment.Center, FontSize = 12, MaxWidth = 78, Margin = new Thickness(0, 3, 0, 0) }; _pendingLabels.Add(lbl); sp.Children.Add(lbl);
        b.Content = sp; return Wire(b, label, run, icon);
    }

    FrameworkElement LargeToggle(string label, string icon, Action run, ToggleButton? register = null)
    {
        var b = new ToggleButton { Padding = new Thickness(5, 4, 5, 2), MinWidth = 62 };
        var sp = new StackPanel();
        sp.Children.Add(Icons.Make(icon, 30));
        var lbl = new TextBlock { Text = label, TextWrapping = TextWrapping.Wrap, TextAlignment = TextAlignment.Center, FontSize = 12, MaxWidth = 78, Margin = new Thickness(0, 3, 0, 0) }; _pendingLabels.Add(lbl); sp.Children.Add(lbl);
        b.Content = sp; b.Tag = label; _transButtons.Add(b);
        return Wire(b, label, run, icon);
    }

    FrameworkElement Small(string label, string icon, Action? run, bool iconOnly = false)
    {
        var b = new Button { Style = (Style)FindResource("RibbonButton"), Padding = new Thickness(5, 2, 8, 2), Height = 26, HorizontalContentAlignment = HorizontalAlignment.Left };
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        sp.Children.Add(Icons.Make(icon, 16));
        if (!iconOnly) { var lbl = new TextBlock { Text = label, FontSize = 12, Margin = new Thickness(6, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center }; _pendingLabels.Add(lbl); sp.Children.Add(lbl); }
        b.Content = sp; return Wire(b, label, run, icon);
    }

    FrameworkElement TextBtn(string text, string tip, Action? run, FontWeight? w = null, FontStyle? st = null, bool underline = false, string? color = null)
    {
        var b = new Button { Style = (Style)FindResource("RibbonButton"), Padding = new Thickness(0), Width = 28, Height = 26, Margin = new Thickness(1) };
        var tb = new TextBlock { Text = text, FontSize = 14, FontWeight = w ?? FontWeights.Normal, FontStyle = st ?? FontStyles.Normal, HorizontalAlignment = HorizontalAlignment.Center };
        if (color != null)
        {
            var c = (Color)ColorConverter.ConvertFromString("#" + color);
            var sp = new StackPanel(); sp.Children.Add(tb);
            sp.Children.Add(new Border { Height = 3, Width = 16, Background = new SolidColorBrush(c), Margin = new Thickness(0, 0, 0, 0) });
            b.Content = sp;
        }
        else { if (underline) tb.TextDecorations = TextDecorations.Underline; tb.SetResourceReference(TextBlock.ForegroundProperty, "FgBrush"); b.Content = tb; }
        return Wire(b, tip, run, "generic");
    }

    static FrameworkElement Stack(params FrameworkElement[] items)
    {
        var s = new StackPanel { VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(2, 2, 2, 0) };
        foreach (var i in items) s.Children.Add(i);
        return s;
    }

    static FrameworkElement Row(params FrameworkElement[] items)
    {
        var s = new StackPanel { Orientation = Orientation.Horizontal };
        foreach (var i in items) s.Children.Add(i);
        return s;
    }

    FrameworkElement Group(string name, params FrameworkElement[] items)
    {
        var g = new Grid();
        g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); g.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        var sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(2, 0, 2, 0) };
        foreach (var i in items) sp.Children.Add(i);
        g.Children.Add(sp);
        var t = new TextBlock { Text = name, FontSize = 11, HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 0, 0, 1) };
        t.SetResourceReference(TextBlock.ForegroundProperty, "MutedBrush"); Grid.SetRow(t, 1); g.Children.Add(t);
        var b = new Border { Child = g, BorderThickness = new Thickness(0, 0, 1, 0), Padding = new Thickness(2, 0, 2, 0) };
        b.SetResourceReference(Border.BorderBrushProperty, "LineBrush");
        var info = new GroupInfo { Element = b, Name = name, Cmds = _pendingCmds.ToList(), Labels = _pendingLabels.ToList(), Pinned = _pendingPinned };
        _pendingCmds.Clear(); _pendingLabels.Clear(); _pendingPinned = false; _groupInfo[b] = info;
        return b;
    }

    TabItem Tab(string header, bool powers, params FrameworkElement[] groups)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        foreach (var g in groups) sp.Children.Add(g);
        var more = new Button { Style = (Style)FindResource("RibbonButton"), Width = 40, Margin = new Thickness(4, 8, 4, 8), Content = Icons.Make("more", 20), Visibility = Visibility.Collapsed, ToolTip = "Altri comandi", VerticalAlignment = VerticalAlignment.Top };
        AutomationProperties.SetName(more, "Altri comandi"); sp.Children.Add(more);
        var t = new TabItem { Header = header, Content = sp, Tag = Res(powers ? "PowersBrush" : "AccentBrush") };
        _tabLayout[t] = (sp, groups.Where(g => _groupInfo.ContainsKey(g)).Select(g => _groupInfo[g]).ToList(), more);
        more.Click += (_, _) => ShowOverflow(t, more);
        AutomationProperties.SetName(t, header);
        if (powers) t.SetResourceReference(Control.ForegroundProperty, "PowersBrush");
        return t;
    }

    FrameworkElement CheckOpt(string label, Func<bool> get, Action<bool> set)
    {
        _pendingPinned = true;   // i gruppi con caselle o interruttori non vanno nel menu a discesa
        var c = new CheckBox { Content = label, FontSize = 12 };
        c.Checked += (_, _) => { if (!_loading) set(true); }; c.Unchecked += (_, _) => { if (!_loading) set(false); };
        AutomationProperties.SetName(c, label); c.Tag = get; _tableChecks.Add(c);
        return c;
    }

    /// <summary>Adatta il ribbon alla larghezza: 1) tutto esteso, 2) gruppi compatti (solo icone), 3) i gruppi in più vanno nel menu "Altri comandi".</summary>
    void AdaptRibbon()
    {
        if (Ribbon.SelectedItem is not TabItem t || !_tabLayout.TryGetValue(t, out var lay)) return;
        double avail = Math.Max(100, Ribbon.ActualWidth - 8 - 8 - 12);
        foreach (var g in lay.Groups) { g.Element.Visibility = Visibility.Visible; foreach (var l in g.Labels) l.Visibility = Visibility.Visible; }
        lay.More.Visibility = Visibility.Collapsed;
        double W() { lay.Panel.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity)); return lay.Panel.DesiredSize.Width; }
        if (W() <= avail) return;
        foreach (var g in lay.Groups) foreach (var l in g.Labels) l.Visibility = Visibility.Collapsed;
        if (W() <= avail) return;
        lay.More.Visibility = Visibility.Visible;
        for (int i = lay.Groups.Count - 1; i >= 0 && W() > avail; i--)
            if (!lay.Groups[i].Pinned) lay.Groups[i].Element.Visibility = Visibility.Collapsed;
    }

    System.Windows.Controls.Primitives.Popup? _overflowPopup;

    void ShowOverflow(TabItem t, Button anchor)
    {
        var hidden = _tabLayout[t].Groups.Where(g => g.Element.Visibility != Visibility.Visible).ToList();
        var sp = new StackPanel { Margin = new Thickness(6) };
        foreach (var g in hidden)
        {
            var h = new TextBlock { Text = g.Name, FontSize = 11, Margin = new Thickness(6, 6, 0, 2) }; h.SetResourceReference(TextBlock.ForegroundProperty, "MutedBrush"); sp.Children.Add(h);
            foreach (var c in g.Cmds)
            {
                var row = new StackPanel { Orientation = Orientation.Horizontal }; row.Children.Add(Icons.Make(c.Icon, 16));
                row.Children.Add(new TextBlock { Text = c.Label, Margin = new Thickness(8, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center, FontSize = 12.5 });
                var b = new Button { Style = (Style)FindResource("RibbonButton"), Content = row, Padding = new Thickness(8, 4, 16, 4), HorizontalContentAlignment = HorizontalAlignment.Left, MinWidth = 200 };
                AutomationProperties.SetName(b, c.Label);
                if (c.Run == null) { b.IsEnabled = false; b.ToolTip = NotYet; }
                else { var run = c.Run; b.Click += (_, _) => { if (_overflowPopup != null) _overflowPopup.IsOpen = false; run(); }; }
                sp.Children.Add(b);
            }
        }
        var border = new Border { BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(8), Child = new ScrollViewer { MaxHeight = 420, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = sp } };
        border.SetResourceReference(Border.BackgroundProperty, "RibbonBrush"); border.SetResourceReference(Border.BorderBrushProperty, "LineBrush");
        _overflowPopup = new System.Windows.Controls.Primitives.Popup { PlacementTarget = anchor, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom, StaysOpen = false, Child = border };
        _overflowPopup.IsOpen = true;
    }

    void BuildRibbon()
    {
        var file = new TabItem { Header = "File", Tag = Res("AccentBrush") }; AutomationProperties.SetName(file, "File");
        Ribbon.Items.Add(file);

        _homeTab = Tab("Home", false,
            Group("Appunti", Large("Incolla", "paste", OnPaste), Stack(Small("Taglia", "cut", OnCut), Small("Copia", "duplicate", OnCopy), Small("Copia formato", "generic", OnCopyFormat))),
            Group("Diapositive", Large("Nuova diapositiva", "slide-new", () => OnNewSlide(this, new RoutedEventArgs())),
                Stack(Small("Duplica", "duplicate", () => OnDuplicate(this, new RoutedEventArgs())), Small("Elimina slide", "trash", () => OnDeleteSlide(this, new RoutedEventArgs())), Small("Nascondi/mostra", "hide", () => OnToggleHidden(this, new RoutedEventArgs())))),
            Group("Carattere", Stack(
                Row(TextBtn("A+", "Aumenta dimensione carattere", () => OnSizeUp(this, new RoutedEventArgs())), TextBtn("A−", "Riduci dimensione carattere", () => OnSizeDown(this, new RoutedEventArgs()))),
                Row(TextBtn("G", "Grassetto", () => OnBold(this, new RoutedEventArgs()), FontWeights.Bold), TextBtn("C", "Corsivo", () => OnItalic(this, new RoutedEventArgs()), null, FontStyles.Italic), TextBtn("S", "Sottolineato", () => OnUnderline(this, new RoutedEventArgs()), null, null, true)),
                Row(TextBtn("A", "Colore carattere arancio", () => FormatColor("ff7a45"), null, null, false, "ff7a45"), TextBtn("A", "Colore carattere blu", () => FormatColor("4f8cff"), null, null, false, "4f8cff"), TextBtn("A", "Colore carattere verde", () => FormatColor("3ecf8e"), null, null, false, "3ecf8e"),
                    TextBtn("A", "Colore carattere rosso", () => FormatColor("c00000"), null, null, false, "c00000"), TextBtn("A", "Colore carattere nero", () => FormatColor("000000"), null, null, false, "000000")))),
            Group("Paragrafo", Stack(
                Row(Small("Elenco puntato", "bullets", () => Bullet("bullet"), true), Small("Elenco numerato", "numbers", () => Bullet("number"), true)),
                Row(Small("Allinea a sinistra", "align-l", () => Align("l"), true), Small("Allinea al centro", "align-c", () => Align("ctr"), true), Small("Allinea a destra", "align-r", () => Align("r"), true), Small("Giustifica", "align-j", () => Align("just"), true)))),
            Group("Disegno", Large("Forme", "shapes", () => OnInsertShape(this, new RoutedEventArgs())), Stack(Small("Elimina forma", "trash", () => OnDeleteShape(this, new RoutedEventArgs())))),
            Group("Modifica", Stack(Small("Trova e sostituisci", "search", OnFindReplace), Small("Seleziona", "generic", OnSelectMenu))));

        var insert = Tab("Inserisci", false,
            Group("Diapositive", Large("Nuova diapositiva", "slide-new", () => OnNewSlide(this, new RoutedEventArgs()))),
            Group("Tabelle", Large("Tabella", "table", () => OnInsertTable(this, new RoutedEventArgs()))),
            Group("Immagini", Large("Immagini", "image", () => OnInsertPicture(this, new RoutedEventArgs()))),
            Group("Illustrazioni", Large("Forme", "shapes", () => OnInsertShape(this, new RoutedEventArgs())), Stack(Small("Icone", "generic", null))),
            Group("Collegamenti", Large("Collegamento", "generic", OnHyperlink)),
            Group("Commenti", Large("Commento", "generic", OnComment)),
            Group("Testo", Large("Casella di testo", "textbox", () => OnInsertText(this, new RoutedEventArgs())), Large("Intestazione e piè di pagina", "generic", OnHeaderFooter)),
            Group("Simboli", Large("Simboli", "generic", OnSymbols)),
            Group("Elementi multimediali", Large("Video", "generic", null), Large("Audio", "generic", null)));

        var trans = Tab("Transizioni", false,
            Group("Anteprima", Large("Anteprima", "preview", () => OnTransPreview())),
            Group("Transizione a questa diapositiva",
                LargeToggle("Nessuna", "tr-none", () => SetTrans("none")), LargeToggle("Dissolvenza", "tr-fade", () => SetTrans("fade")), LargeToggle("Spingi", "tr-push", () => SetTrans("push")),
                LargeToggle("Cancella", "tr-wipe", () => SetTrans("wipe")), LargeToggle("Copri", "tr-cover", () => SetTrans("cover")), LargeToggle("Scopri", "tr-pull", () => SetTrans("pull"))),
            Group("Opzioni effetto", Stack(Row(Small("Verso sinistra", "dir-l", () => SetTransDir("l"), true), Small("Verso destra", "dir-r", () => SetTransDir("r"), true)), Row(Small("Verso l'alto", "dir-u", () => SetTransDir("u"), true), Small("Verso il basso", "dir-d", () => SetTransDir("d"), true)))),
            Group("Intervallo", Stack(Small("Durata lenta", "slow", () => SetTransSpeed("slow")), Small("Durata media", "med", () => SetTransSpeed("med")), Small("Durata veloce", "fast", () => SetTransSpeed("fast"))), Stack(Small("Applica a tutte", "duplicate", OnApplyTransAll))));

        var show = Tab("Presentazione", false,
            Group("Avvia presentazione", Large("Dall'inizio", "play", () => OnPresent(this, new RoutedEventArgs())), Large("Dalla diapositiva corrente", "play-here", () => OnPresentHere(this, new RoutedEventArgs())), Large("Vista relatore", "presenter", () => OnPresenter(this, new RoutedEventArgs())), Large("Presentazione personalizzata", "generic", null)),
            Group("Imposta", Large("Imposta presentazione", "generic", OnSetupShow), Large("Nascondi diapositiva", "hide", () => OnToggleHidden(this, new RoutedEventArgs())), Large("Prova intervalli", "generic", null), Large("Registra", "generic", null)));

        var view = Tab("Visualizza", false,
            Group("Visualizzazioni presentazione", Large("Normale", "normal", () => SetSorter(false)), Large("Sequenza diapositive", "sorter", () => SetSorter(true)), Large("Pagina note", "notes", null)),
            Group("Mostra", Stack(CheckOpt("Note", () => _showNotes, v => SetNotesVisible(v)), CheckOpt("Righello", () => false, _ => { }), CheckOpt("Griglia", () => false, _ => { }))),
            Group("Zoom", Large("Adatta alla finestra", "fit", () => SetZoom(100))));
        foreach (var c in _tableChecks.Where(c => (string?)AutomationProperties.GetName(c) is "Righello" or "Griglia")) { c.IsEnabled = false; c.ToolTip = NotYet; }

        _darkSwitch = new ToggleButton { Style = (Style)FindResource("SwitchToggle"), IsChecked = _settings.Dark, Margin = new Thickness(8, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        AutomationProperties.SetName(_darkSwitch, "Modalità scura");
        _darkSwitch.Checked += (_, _) => SetDark(true); _darkSwitch.Unchecked += (_, _) => SetDark(false);
        _pendingPinned = true;
        var darkRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(10, 14, 10, 0), VerticalAlignment = VerticalAlignment.Top };
        darkRow.Children.Add(Icons.Make("moon", 22)); darkRow.Children.Add(new TextBlock { Text = "Modalità scura", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0, 0, 0), FontSize = 13 }); darkRow.Children.Add(_darkSwitch);
        var powers = Tab("Powers", true, Group("Aspetto", darkRow), Group("Privacy", Large("Scudo privacy", "shield", OnPrivacy)));

        _notesCheck = _tableChecks.First(c => AutomationProperties.GetName(c) == "Note");
        _tableChecks.Clear();
        _tableTab = Tab("Struttura tabella", false,
            Group("Opzioni stile tabella", Stack(
                CheckOpt("Riga di intestazione", () => false, v => SetTableFlag("firstRow", v)), CheckOpt("Riga Totale", () => false, v => SetTableFlag("lastRow", v)), CheckOpt("Righe alternate evidenziate", () => false, v => SetTableFlag("bandRow", v))),
                Stack(CheckOpt("Prima colonna", () => false, v => SetTableFlag("firstCol", v)), CheckOpt("Ultima colonna", () => false, v => SetTableFlag("lastCol", v)), CheckOpt("Colonne alternate evidenziate", () => false, v => SetTableFlag("bandCol", v)))));
        _tableTab.Visibility = Visibility.Collapsed; _tableTab.SetResourceReference(Control.ForegroundProperty, "AccentBrush");

        foreach (var t in new[] { _homeTab, insert, trans, show, view, powers, _tableTab }) Ribbon.Items.Add(t);
        Ribbon.SelectedItem = _homeTab;
        Ribbon.SizeChanged += (_, _) => { AdaptRibbon(); AdaptChrome(); };
        Ribbon.SelectionChanged += (_, _) => Dispatcher.BeginInvoke(AdaptRibbon, System.Windows.Threading.DispatcherPriority.Loaded);
    }
}
