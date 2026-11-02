using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using PowerPresentations.Core;

namespace PowerPresentations.App;

/// <summary>Trova e sostituisci, collegamento, simboli, copia formato.</summary>
public partial class MainWindow
{
    string _findText = "", _replText = ""; bool _matchCase;
    FormatClip? _fmtClip; int _fmtSrc;

    static (int Button, string Text) Prompt(Window owner, string title, string message, string initial, Func<string, string?> check, params string[] extra)
    {
        var tb = new TextBox { Text = initial, FontSize = 13.5, Padding = new Thickness(6, 5, 6, 5), Margin = new Thickness(0, 14, 0, 0) }; AutomationProperties.SetName(tb, title);
        var err = new TextBlock { Foreground = (System.Windows.Media.Brush)Application.Current.Resources["PowersBrush"], FontSize = 12, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 8, 0, 0), Visibility = Visibility.Collapsed };
        var sp = new StackPanel(); sp.Children.Add(tb); sp.Children.Add(err);
        var d = new ThemedDialog(owner, title, message, DialogKind.Question, sp, true);
        d.AddButton("Annulla", false, true);
        foreach (var x in extra) d.AddButton(x);
        d.AddButton("OK", true, false, () => { var e = check(tb.Text); err.Text = e ?? ""; err.Visibility = e == null ? Visibility.Collapsed : Visibility.Visible; return e == null; });
        d.Loaded += (_, _) => { tb.Focus(); tb.SelectAll(); };
        d.ShowDialog();
        int ok = 1 + extra.Length;
        return (d.Result == ok ? 0 : d.Result <= 0 ? -1 : d.Result, tb.Text);   // 0 = OK, -1 = annulla, n = pulsante extra n
    }

    // ---------- trova e sostituisci ----------
    void OnFindReplace()
    {
        if (_doc == null) return;
        CommitEdit();
        var f = new TextBox { Text = _findText, FontSize = 13.5, Padding = new Thickness(6, 5, 6, 5) }; AutomationProperties.SetName(f, "Trova");
        var r = new TextBox { Text = _replText, FontSize = 13.5, Padding = new Thickness(6, 5, 6, 5) }; AutomationProperties.SetName(r, "Sostituisci con");
        var cs = new CheckBox { Content = "Maiuscole/minuscole", IsChecked = _matchCase, Margin = new Thickness(0, 10, 0, 0) }; AutomationProperties.SetName(cs, "Maiuscole/minuscole");
        var sp = new StackPanel { Margin = new Thickness(0, 12, 0, 0) };
        TextBlock L(string t) => new() { Text = t, FontSize = 12, Margin = new Thickness(0, 6, 0, 4), Foreground = (System.Windows.Media.Brush)Application.Current.Resources["MutedBrush"] };
        sp.Children.Add(L("Trova")); sp.Children.Add(f); sp.Children.Add(L("Sostituisci con")); sp.Children.Add(r); sp.Children.Add(cs);
        var d = new ThemedDialog(this, "Trova e sostituisci", "Cerca nel testo di tutte le slide, tabelle comprese.", DialogKind.Question, sp, true);
        d.AddButton("Chiudi", false, true);
        d.AddButton("Trova successivo", false, false, () => f.Text.Length > 0);
        d.AddButton("Sostituisci tutto", true, false, () => f.Text.Length > 0);
        d.Loaded += (_, _) => { f.Focus(); f.SelectAll(); };
        d.ShowDialog();
        _findText = f.Text; _replText = r.Text; _matchCase = cs.IsChecked == true;
        if (_findText.Length == 0) return;
        if (d.Result == 1) FindNext(); else if (d.Result == 2) ReplaceEverywhere();
    }

    void FindNext()
    {
        var hits = _doc!.FindAll(_findText, _matchCase);
        if (hits.Count == 0) { Status.Text = $"«{_findText}» non trovato."; return; }
        int Key(int slide, int id) => slide * 100000 + _doc.ShapeIds(slide).IndexOf(id);
        int cur = Key(_cur, View.SelectedId);
        var next = hits.Select(h => (h, k: Key(h.Slide, h.ShapeId))).Where(x => x.k > cur).OrderBy(x => x.k).Select(x => x.h).FirstOrDefault();
        bool wrapped = next == null; next ??= hits.OrderBy(h => Key(h.Slide, h.ShapeId)).First();
        SetSorter(false); ShowSlide(next.Slide); View.SelectedId = next.ShapeId; OnSelect(next.ShapeId); View.Focus();
        Status.Text = $"Trovato nella diapositiva {next.Slide + 1}" + (wrapped ? " (ricerca ripartita dall'inizio)" : "");
    }

    void ReplaceEverywhere()
    {
        if (_doc!.FindAll(_findText, _matchCase).Count == 0) { Status.Text = $"«{_findText}» non trovato."; return; }
        int n = 0; SaveNotes(); Snap(); n = _doc.ReplaceAll(_findText, _replText, _matchCase);
        RebuildThumbs(); ShowSlide(_cur); UpdateTitle(); Status.Text = n == 1 ? "1 sostituzione." : $"{n} sostituzioni.";
    }

    // ---------- collegamento ----------
    void OnHyperlink()
    {
        int id = View.SelectedId;
        if (_doc == null || id == 0 || _doc.TextBodies(_cur, id).Count == 0) { Status.Text = "Seleziona prima una forma con testo da collegare."; return; }
        var old = _doc.GetHyperlink(_cur, id);
        var (b, text) = Prompt(this, "Collegamento", "Indirizzo (http, https o mailto). Il collegamento si applica a tutto il testo della forma.", old ?? "https://",
            t => { try { PptxDocument.ValidateLink(t); return null; } catch (ArgumentException e) { return e.Message; } }, old != null ? new[] { "Rimuovi" } : Array.Empty<string>());
        if (b < 0) return;
        Snap(); _doc.SetHyperlink(_cur, id, b == 0 ? text : null); RefreshCurrent(); View.SelectedId = id;
    }

    // ---------- simboli ----------
    static readonly (string Ch, string Name)[] Symbols =
    {
        ("©", "Copyright"), ("®", "Marchio registrato"), ("™", "Marchio"), ("€", "Euro"), ("£", "Sterlina"), ("¥", "Yen"), ("§", "Paragrafo"), ("°", "Gradi"), ("±", "Più o meno"),
        ("×", "Per"), ("÷", "Diviso"), ("≠", "Diverso"), ("≤", "Minore o uguale"), ("≥", "Maggiore o uguale"), ("∞", "Infinito"), ("π", "Pi greco"), ("Ω", "Omega"), ("µ", "Micro"),
        ("‰", "Per mille"), ("→", "Freccia destra"), ("←", "Freccia sinistra"), ("↑", "Freccia su"), ("↓", "Freccia giù"), ("✓", "Segno di spunta"), ("★", "Stella"), ("…", "Puntini di sospensione"), ("–", "Trattino en"), ("—", "Trattino em")
    };

    void OnSymbols()
    {
        if (_doc == null) return;
        int k = ListPicker.Pick(this, "Simboli", Symbols.Select(s => $"{s.Ch}   {s.Name}")); if (k < 0) return;
        var ch = Symbols[k].Ch; int id = View.SelectedId;
        if (id != 0 && _doc.TextBodies(_cur, id).Count > 0) { Guard(() => _doc.AppendText(_cur, id, ch)); RefreshCurrent(); View.SelectedId = id; }
        else
        {
            var (w, h) = _doc.SlideSize; int nid = 0;
            Guard(() => nid = _doc.AddTextBox(_cur, w / 2 - 50, h / 2 - 20, 100, 40, ch)); RefreshCurrent(); View.SelectedId = nid;
        }
    }

    // ---------- copia formato ----------
    void OnCopyFormat()
    {
        int id = View.SelectedId;
        if (_doc == null || id == 0) { Status.Text = "Seleziona prima la forma da cui copiare il formato."; return; }
        try { _fmtClip = _doc.CopyFormat(_cur, id); _fmtSrc = id; Status.Text = "Copia formato: fai clic sulla forma di destinazione."; }
        catch (KeyNotFoundException) { Status.Text = "Formato non copiabile da questo oggetto."; }
    }

    /// <summary>Chiamata a ogni clic su una forma: se il pennello è armato applica il formato.</summary>
    void ApplyArmedFormat(int id)
    {
        var fc = _fmtClip; _fmtClip = null;
        if (fc == null || _doc == null || id == 0 || id == _fmtSrc) return;
        if (Try(() => _doc.ApplyFormat(_cur, id, fc), "Copia formato non riuscita")) { RefreshCurrent(); View.SelectedId = id; Status.Text = "Formato applicato."; }
    }
}
