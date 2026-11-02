using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using PowerPresentations.Core;

namespace PowerPresentations.App;

/// <summary>Commento, Intestazione e piè di pagina, Imposta presentazione.</summary>
public partial class MainWindow
{
    static Brush Br(string k) => (Brush)Application.Current.Resources[k];

    static CheckBox Check(string text, bool on, Thickness? m = null) { var c = new CheckBox { Content = text, IsChecked = on, Margin = m ?? new Thickness(0, 8, 0, 0), Foreground = Br("FgBrush") }; AutomationProperties.SetName(c, text); return c; }
    static RadioButton Radio(string text, bool on, string group, Thickness? m = null) { var c = new RadioButton { Content = text, IsChecked = on, GroupName = group, Margin = m ?? new Thickness(0, 6, 0, 0), Foreground = Br("FgBrush") }; AutomationProperties.SetName(c, text); return c; }
    static TextBox Field(string name, string text) { var t = new TextBox { Text = text, FontSize = 13, Padding = new Thickness(6, 4, 6, 4), Margin = new Thickness(0, 6, 0, 0) }; AutomationProperties.SetName(t, name); return t; }
    static TextBlock Caption(string t) => new() { Text = t, FontSize = 12, Foreground = Br("MutedBrush"), Margin = new Thickness(0, 10, 0, 0) };

    // ---------- commento ----------
    void OnComment()
    {
        if (_doc == null) return;
        CommitEdit();
        var sp = new StackPanel { Margin = new Thickness(0, 10, 0, 0) };
        var existing = _doc.GetComments(_cur);
        if (existing.Count > 0)
        {
            sp.Children.Add(Caption($"Commenti già presenti: {existing.Count}"));
            sp.Children.Add(new TextBlock { Text = string.Join("\n", existing.Select(c => $"{c.Author}: {c.Text}")), TextWrapping = TextWrapping.Wrap, MaxHeight = 120, FontSize = 12.5, Margin = new Thickness(0, 4, 0, 0), Foreground = Br("FgBrush") });
        }
        sp.Children.Add(Caption("Nuovo commento"));
        var tb = new TextBox { AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, MinHeight = 70, MaxLength = 10000, FontSize = 13, Padding = new Thickness(6, 4, 6, 4), Margin = new Thickness(0, 4, 0, 0) }; AutomationProperties.SetName(tb, "Testo del commento");
        sp.Children.Add(tb);
        var err = new TextBlock { Foreground = Br("PowersBrush"), FontSize = 12, Visibility = Visibility.Collapsed, Margin = new Thickness(0, 6, 0, 0) }; sp.Children.Add(err);
        var d = new ThemedDialog(this, "Commento", "Il commento si aggiunge alla diapositiva " + (_cur + 1) + ".", DialogKind.Question, sp, true);
        d.AddButton("Chiudi", false, true);
        d.AddButton("Aggiungi", true, false, () => { if (tb.Text.Trim().Length > 0) return true; err.Text = "Scrivi il testo del commento."; err.Visibility = Visibility.Visible; return false; });
        d.Loaded += (_, _) => tb.Focus();
        d.ShowDialog();
        if (d.Result != 1) return;
        var text = tb.Text;
        if (Try(() => _doc.AddComment(_cur, text), "Commento non aggiunto")) { UpdateTitle(); Status.Text = $"Commento aggiunto ({_doc.GetComments(_cur).Count} nella diapositiva)."; }
    }

    // ---------- intestazione e piè di pagina ----------
    void OnHeaderFooter()
    {
        if (_doc == null) return;
        CommitEdit();
        var cur = _doc.GetHeaderFooter(_cur);
        var sp = new StackPanel { Margin = new Thickness(0, 10, 0, 0) };
        var dt = Check("Data e ora", cur.Date, new Thickness(0)); var auto = Radio("Aggiorna automaticamente", cur.DateAuto || !cur.Date, "dt", new Thickness(24, 4, 0, 0)); var fix = Radio("Fissa", cur.Date && !cur.DateAuto, "dt", new Thickness(24, 4, 0, 0));
        var dtext = Field("Data fissa", cur.DateText); dtext.Margin = new Thickness(24, 4, 0, 0);
        var num = Check("Numero diapositiva", cur.SlideNumber); var ft = Check("Piè di pagina", cur.Footer); var ftext = Field("Testo del piè di pagina", cur.FooterText); ftext.MaxLength = 500; ftext.Margin = new Thickness(24, 4, 0, 0);
        var skip = Check("Non mostrare nella diapositiva del titolo", false);
        foreach (var e in new UIElement[] { dt, auto, fix, dtext, num, ft, ftext, skip }) sp.Children.Add(e);
        var d = new ThemedDialog(this, "Intestazione e piè di pagina", "Il segnaposto compare solo se il layout della diapositiva lo prevede.", DialogKind.Question, sp, true);
        d.AddButton("Annulla", false, true); d.AddButton("Applica"); d.AddButton("Applica a tutte", true);
        d.ShowDialog();
        if (d.Result is not (1 or 2)) return;
        var hf = new HeaderFooter(num.IsChecked == true, dt.IsChecked == true, fix.IsChecked != true, dtext.Text, ft.IsChecked == true, ftext.Text, skip.IsChecked == true);
        int target = d.Result == 1 ? _cur : -1;
        SaveNotes();
        if (Try(() => _doc.ApplyHeaderFooter(hf, target), "Intestazione e piè di pagina")) { RebuildThumbs(); ShowSlide(_cur); UpdateTitle(); }
    }

    // ---------- imposta presentazione ----------
    void OnSetupShow()
    {
        if (_doc == null) return;
        var s = _doc.GetShowSettings(); int n = _doc.Slides.Count;
        var sp = new StackPanel { Margin = new Thickness(0, 10, 0, 0) };
        var pres = Radio("Presentata da un relatore (schermo intero)", s.Mode == "present", "mode", new Thickness(0)); var brow = Radio("Sfogliata da una persona (finestra)", s.Mode == "browse", "mode"); var kiosk = Radio("Sfogliata in postazione informatica (schermo intero)", s.Mode == "kiosk", "mode");
        var loop = Check("Ripeti in ciclo continuo", s.Loop); var nar = Check("Senza narrazione", !s.Narration); var ani = Check("Senza animazioni", !s.Animation); var tim = Check("Usa gli intervalli salvati", s.UseTimings);
        var all = Radio("Tutte le diapositive", s.From == null, "range", new Thickness(0, 12, 0, 0)); var rg = Radio("Da diapositiva a diapositiva", s.From != null, "range");
        var from = Field("Da", (s.From ?? 1).ToString()); var to = Field("A", (s.To ?? n).ToString()); from.Width = to.Width = 70; from.HorizontalAlignment = to.HorizontalAlignment = HorizontalAlignment.Left;
        var err = new TextBlock { Foreground = Br("PowersBrush"), FontSize = 12, TextWrapping = TextWrapping.Wrap, Visibility = Visibility.Collapsed, Margin = new Thickness(0, 8, 0, 0) };
        foreach (var e in new UIElement[] { pres, brow, kiosk, loop, nar, ani, tim, all, rg, from, to, err }) sp.Children.Add(e);
        ShowSettings? result = null;
        var d = new ThemedDialog(this, "Imposta presentazione", "Come viene riprodotta la presentazione.", DialogKind.Question, sp, true);
        d.AddButton("Annulla", false, true);
        d.AddButton("OK", true, false, () =>
        {
            int? f = null, t = null;
            if (rg.IsChecked == true)
            {
                if (!int.TryParse(from.Text, out var a) || !int.TryParse(to.Text, out var b) || a < 1 || b > n || a > b) { err.Text = $"Intervallo non valido: numeri da 1 a {n}, il primo non oltre il secondo."; err.Visibility = Visibility.Visible; return false; }
                f = a; t = b;
            }
            result = new ShowSettings(kiosk.IsChecked == true ? "kiosk" : brow.IsChecked == true ? "browse" : "present", loop.IsChecked == true, nar.IsChecked != true, ani.IsChecked != true, tim.IsChecked == true, f, t);
            return true;
        });
        d.ShowDialog();
        if (result == null || result == s) return;
        var res = result;
        Try(() => _doc.SetShowSettings(res), "Imposta presentazione");
        UpdateTitle();
    }
}
