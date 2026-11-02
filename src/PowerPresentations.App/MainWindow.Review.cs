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
    /// <summary>Il comando Commento apre/chiude il riquadro Commenti (elenco della diapositiva corrente).</summary>
    void OnComment()
    {
        if (_doc == null) return;
        CommitEdit();
        CommentsPane.Visibility = CommentsPane.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
        RefreshComments();
    }

    void OnCommentsClose(object s, RoutedEventArgs e) { CommentsPane.Visibility = Visibility.Collapsed; View.Focus(); }

    void OnNewCommentClick(object s, RoutedEventArgs e) => OnNewComment();

    /// <summary>Ricostruisce l'elenco; da chiamare quando cambia slide o dopo ogni operazione sui commenti.</summary>
    void RefreshComments()
    {
        if (_doc == null || CommentsPane.Visibility != Visibility.Visible) return;
        CommentList.Children.Clear();
        var list = _doc.GetComments(_cur);
        CommentsTitle.Text = $"Commenti ({list.Count})";
        if (list.Count == 0) CommentList.Children.Add(new TextBlock { Text = "Nessun commento in questa diapositiva.", Foreground = Br("MutedBrush"), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(6, 8, 6, 0), FontSize = 12.5 });
        int n = 0;
        foreach (var c in list)
        {
            n++; var ci = c; var panel = new StackPanel();
            panel.Children.Add(new TextBlock { Text = (c.IsReply ? "Risposta di " : "") + c.Author + (CommentDate(c.Created) is { Length: > 0 } dt ? " · " + dt : ""), FontSize = 11.5, Foreground = Br("MutedBrush") });
            panel.Children.Add(new TextBlock { Text = c.Text, TextWrapping = TextWrapping.Wrap, FontSize = 13, Foreground = Br("FgBrush"), Margin = new Thickness(0, 3, 0, 0) });
            var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 0) };
            var edit = new Button { Content = "Modifica", Padding = new Thickness(10, 2, 10, 2) }; AutomationProperties.SetName(edit, "Modifica commento " + n);
            var del = new Button { Content = "Elimina", Padding = new Thickness(10, 2, 10, 2), Margin = new Thickness(6, 0, 0, 0) }; AutomationProperties.SetName(del, "Elimina commento " + n);
            edit.Click += (_, _) => EditComment(ci); del.Click += (_, _) => DeleteComment(ci);
            row.Children.Add(edit); row.Children.Add(del); panel.Children.Add(row);
            CommentList.Children.Add(new Border { Child = panel, Padding = new Thickness(8), Margin = new Thickness(c.IsReply ? 22 : 0, 0, 0, 6), BorderBrush = Br("LineBrush"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6) });
        }
    }

    static string CommentDate(string iso) => System.DateTime.TryParse(iso, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.RoundtripKind, out var d) ? d.ToString("g", System.Globalization.CultureInfo.CurrentCulture) : iso;

    void EditComment(CommentInfo c)
    {
        CommitEdit();
        var sp = new StackPanel { Margin = new Thickness(0, 10, 0, 0) };
        var tb = new TextBox { Text = c.Text, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, MinHeight = 70, MaxLength = 10000, FontSize = 13, Padding = new Thickness(6, 4, 6, 4) }; AutomationProperties.SetName(tb, "Testo del commento");
        var err = new TextBlock { Foreground = Br("PowersBrush"), FontSize = 12, Visibility = Visibility.Collapsed, Margin = new Thickness(0, 6, 0, 0) };
        sp.Children.Add(tb); sp.Children.Add(err);
        var d = new ThemedDialog(this, "Modifica commento", "Autore e data originali restano invariati.", DialogKind.Question, sp, true);
        d.AddButton("Annulla", false, true);
        d.AddButton("Salva", true, false, () => { if (tb.Text.Trim().Length > 0) return true; err.Text = "Scrivi il testo del commento."; err.Visibility = Visibility.Visible; return false; });
        d.Loaded += (_, _) => { tb.Focus(); tb.SelectAll(); };
        d.ShowDialog();
        if (d.Result != 1) return;
        var text = tb.Text;
        if (Try(() => _doc!.EditComment(_cur, c.Id, text), "Commento non modificato")) { UpdateTitle(); Status.Text = "Commento modificato."; }
        RefreshComments();
    }

    void DeleteComment(CommentInfo c)
    {
        CommitEdit();
        if (c.Replies > 0 && !ThemedDialog.Confirm(this, "Elimina commento", $"Il commento ha {c.Replies} {(c.Replies == 1 ? "risposta, che verrà eliminata" : "risposte, che verranno eliminate")} insieme a lui. Continuare?", "Elimina")) return;
        if (Try(() => _doc!.DeleteComment(_cur, c.Id), "Commento non eliminato")) { UpdateTitle(); Status.Text = "Commento eliminato."; }
        RefreshComments();
    }

    void OnNewComment()
    {
        if (_doc == null) return;
        CommitEdit();
        var sp = new StackPanel { Margin = new Thickness(0, 10, 0, 0) };
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
        RefreshComments();
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
