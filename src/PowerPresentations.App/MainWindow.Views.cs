using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using PowerPresentations.Core;

namespace PowerPresentations.App;

/// <summary>Prova intervalli e Pagina note.</summary>
public partial class MainWindow
{
    // ---------- prova intervalli ----------
    void OnRehearse()
    {
        if (_doc == null) return;
        CommitEdit(); SaveNotes();
        var c = new ShowController(_doc, _r!, 0) { Rehearse = true };
        var show = new PresentationWindow(c);
        show.Closed += (_, _) => FinishRehearsal(c);
        show.ShowFullScreen(null);
    }

    void FinishRehearsal(ShowController c)
    {
        var times = c.Spent.Where(kv => kv.Value.TotalSeconds >= 0.5).OrderBy(kv => kv.Key).ToList();
        Activate();
        if (_doc == null || times.Count == 0) return;
        var total = TimeSpan.FromTicks(times.Sum(kv => kv.Value.Ticks));
        if (!ThemedDialog.Confirm(this, "Prova intervalli", $"Durata totale: {total.ToString(@"hh\:mm\:ss")}. Vuoi conservare i nuovi intervalli per {times.Count} diapositive?", "Conserva", "Scarta")) return;
        Snap();
        foreach (var (slide, ts) in times) if (slide < _doc.Slides.Count) _doc.SetAdvanceTime(slide, (int)(Math.Round(ts.TotalMilliseconds / 100) * 100));
        UpdateTitle(); Status.Text = "Intervalli salvati: la presentazione avanza da sola se \"Usa gli intervalli salvati\" è attivo.";
    }

    // ---------- pagina note ----------
    void OnNotesPage()
    {
        if (_doc == null) return;
        CommitEdit(); SaveNotes();
        int idx = _cur;
        var img = new Image { Stretch = Stretch.Uniform, Margin = new Thickness(0, 0, 0, 10), MaxHeight = 340 }; AutomationProperties.SetName(img, "Immagine della diapositiva");
        var tb = new TextBox { AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, MinHeight = 160, FontSize = 14, Padding = new Thickness(8), VerticalScrollBarVisibility = ScrollBarVisibility.Auto }; AutomationProperties.SetName(tb, "Note della pagina");
        var label = new TextBlock { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(12, 0, 12, 0), Foreground = Br("FgBrush") };
        var prev = new Button { Content = "◀ Precedente", Padding = new Thickness(12, 4, 12, 4) }; var next = new Button { Content = "Successivo ▶", Padding = new Thickness(12, 4, 12, 4), Margin = new Thickness(6, 0, 0, 0) }; var close = new Button { Content = "Chiudi", Padding = new Thickness(14, 4, 14, 4), Margin = new Thickness(18, 0, 0, 0), IsCancel = true };
        AutomationProperties.SetName(prev, "Precedente"); AutomationProperties.SetName(next, "Successivo"); AutomationProperties.SetName(close, "Chiudi pagina note");
        var bar = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 0) };
        foreach (var e in new UIElement[] { prev, next, label, close }) bar.Children.Add(e);
        var root = new DockPanel { Margin = new Thickness(18) };
        DockPanel.SetDock(bar, Dock.Bottom); DockPanel.SetDock(img, Dock.Top);
        root.Children.Add(bar); root.Children.Add(img); root.Children.Add(tb);
        var w = new Window { Title = "Pagina note", Owner = this, Width = 760, Height = 700, MinWidth = 520, WindowStartupLocation = WindowStartupLocation.CenterOwner, Background = Br("BgBrush"), Foreground = Br("FgBrush"), Content = root, ShowInTaskbar = false };
        AutomationProperties.SetName(w, "Pagina note");
        void Load() { img.Source = SlideDrawing.ToBitmap(_r!.Render(idx), 900); tb.Text = _doc.GetNotes(idx); label.Text = $"Diapositiva {idx + 1} di {_doc.Slides.Count}"; prev.IsEnabled = idx > 0; next.IsEnabled = idx < _doc.Slides.Count - 1; tb.CaretIndex = tb.Text.Length; }
        void Commit() { if (tb.Text != _doc.GetNotes(idx)) { Snap(); _doc.SetNotes(idx, tb.Text); UpdateTitle(); } }
        prev.Click += (_, _) => { Commit(); idx--; Load(); }; next.Click += (_, _) => { Commit(); idx++; Load(); };
        w.Closing += (_, _) => Commit();
        w.Loaded += (_, _) => { Load(); tb.Focus(); };
        w.ShowDialog();
        RebuildThumbs(); ShowSlide(idx);
    }

    // ---------- stampa ----------
    void OnPrint()
    {
        if (_doc == null) return;
        CommitEdit(); SaveNotes();
        var pd = new PrintDialog { UserPageRangeEnabled = false };
        pd.PrintTicket.PageOrientation = System.Printing.PageOrientation.Landscape;
        if (pd.ShowDialog() != true) return;
        try
        {
            var fd = SlidePrint.Build(_doc, _r!, pd.PrintableAreaWidth, pd.PrintableAreaHeight);
            if (fd.Pages.Count == 0) { ThemedDialog.Notify(this, "Stampa", "Tutte le diapositive sono nascoste: niente da stampare.", DialogKind.Info); return; }
            pd.PrintDocument(fd.DocumentPaginator, Title);
            Status.Text = $"Inviate alla stampante {fd.Pages.Count} pagine.";
        }
        catch (Exception ex) when (ex is System.Printing.PrintQueueException or System.Printing.PrintSystemException or InvalidOperationException or System.IO.IOException)
        { ThemedDialog.Notify(this, "Stampa non riuscita", ex.Message, DialogKind.Error); }
    }
}
