using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using PowerPresentations.Core;

namespace PowerPresentations.App;

/// <summary>Presentazione personalizzata: elenco, creazione/modifica, eliminazione, avvio.</summary>
public partial class MainWindow
{
    void OnCustomShows()
    {
        if (_doc == null) return;
        CommitEdit();
        var shows = _doc.CustomShows;
        var items = new List<string> { "Nuova presentazione personalizzata…" }; items.AddRange(shows.Select(s => $"{s.Name}  ({s.Slides.Count} diapositive)"));
        int k = ListPicker.Pick(this, "Presentazione personalizzata", items); if (k < 0) return;
        if (k == 0) { EditCustomShow(null); return; }
        var show = shows[k - 1];
        var d = new ThemedDialog(this, show.Name, $"{show.Slides.Count} diapositive: " + string.Join(", ", show.Slides.Select(i => i + 1)), DialogKind.Question, null, true);
        d.AddButton("Annulla", false, true); d.AddButton("Elimina"); d.AddButton("Modifica"); d.AddButton("Avvia", true);
        d.ShowDialog();
        switch (d.Result)
        {
            case 1: if (Try(() => _doc.DeleteCustomShow(show.Name), "Presentazione personalizzata")) UpdateTitle(); break;
            case 2: EditCustomShow(show); break;
            case 3: StartShow(0, false, true, show.Slides); break;
        }
    }

    void EditCustomShow(CustomShowInfo? existing)
    {
        int n = _doc!.Slides.Count;
        var name = Field("Nome", existing?.Name ?? $"Presentazione personalizzata {_doc.CustomShows.Count + 1}"); name.MaxLength = 100;
        var list = new StackPanel();
        var boxes = new List<CheckBox>();
        for (int i = 0; i < n; i++)
        {
            var first = _r!.Render(i).Items.OfType<PowerPresentations.Core.Rendering.TextItem>().FirstOrDefault()?.Text ?? "";
            var c = Check($"{i + 1}. {(first.Length > 50 ? first[..50] + "…" : first)}", existing?.Slides.Contains(i) ?? true, new Thickness(0, 3, 0, 0)); boxes.Add(c); list.Children.Add(c);
        }
        var err = new TextBlock { Foreground = Br("PowersBrush"), FontSize = 12, Visibility = Visibility.Collapsed, Margin = new Thickness(0, 6, 0, 0) };
        var sp = new StackPanel { Margin = new Thickness(0, 8, 0, 0) };
        sp.Children.Add(Caption("Nome")); sp.Children.Add(name); sp.Children.Add(Caption("Diapositive incluse (nell'ordine del documento)"));
        sp.Children.Add(new ScrollViewer { Content = list, MaxHeight = 220, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Margin = new Thickness(0, 4, 0, 0) }); sp.Children.Add(err);
        var d = new ThemedDialog(this, existing == null ? "Nuova presentazione personalizzata" : "Modifica presentazione personalizzata", "Scegli le diapositive da includere.", DialogKind.Question, sp, true);
        d.AddButton("Annulla", false, true);
        d.AddButton("OK", true, false, () =>
        {
            if (name.Text.Trim().Length == 0) { err.Text = "Il nome è obbligatorio."; err.Visibility = Visibility.Visible; return false; }
            if (!boxes.Any(b => b.IsChecked == true)) { err.Text = "Scegli almeno una diapositiva."; err.Visibility = Visibility.Visible; return false; }
            return true;
        });
        d.ShowDialog();
        if (d.Result != 1) return;
        var picked = boxes.Select((b, i) => (b, i)).Where(x => x.b.IsChecked == true).Select(x => x.i).ToList(); var nm = name.Text;
        // se si rinomina, la vecchia voce esce
        Try(() => { if (existing != null && !string.Equals(existing.Name, nm.Trim(), StringComparison.CurrentCultureIgnoreCase)) _doc.DeleteCustomShow(existing.Name); _doc.SetCustomShow(nm, picked); }, "Presentazione personalizzata");
        UpdateTitle();
    }
}
