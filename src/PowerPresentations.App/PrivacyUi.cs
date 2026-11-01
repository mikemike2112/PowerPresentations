using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using PowerPresentations.Core;

namespace PowerPresentations.App;

/// <summary>Scudo privacy: analizza il documento aperto e crea una COPIA ripulita; il file aperto e l'originale non vengono mai toccati.</summary>
public partial class MainWindow
{
    ThemedDialog BuildPrivacyDialog(PrivacyReport rep, out Func<(HashSet<string> Ids, CommentMode Mode)> choice)
    {
        var panel = new StackPanel { Margin = new Thickness(0, 12, 0, 0) };
        var boxes = new List<(CheckBox Box, string Id)>();
        TextBlock H(string t) => new() { Text = t, FontWeight = FontWeights.SemiBold, FontSize = 12.5, Margin = new Thickness(0, 10, 0, 4), Foreground = (Brush)FindResource("FgBrush") };
        RadioButton? anon = null, remove = null;
        if (rep.Findings.Count == 0)
            panel.Children.Add(new TextBlock { Text = "Tra i dati che so cercare non ho trovato nulla di personale.", TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 4, 0, 0), Foreground = (Brush)FindResource("FgBrush") });
        foreach (var g in rep.Findings.GroupBy(f => f.Category))
        {
            panel.Children.Add(H(g.Key));
            foreach (var f in g)
            {
                var cb = new CheckBox { IsChecked = true, Margin = new Thickness(2, 2, 0, 2), Content = new TextBlock { Text = $"{f.Label}: {f.Value}", TextWrapping = TextWrapping.Wrap, Foreground = (Brush)FindResource("FgBrush") } };
                AutomationProperties.SetName(cb, $"{f.Label}: {f.Value}"); boxes.Add((cb, f.Id)); panel.Children.Add(cb);
                if (f.Id == "comments")
                {
                    anon = new RadioButton { Content = "Anonimizza gli autori (contenuto invariato)", IsChecked = true, GroupName = "cm", Margin = new Thickness(26, 2, 0, 0), Foreground = (Brush)FindResource("FgBrush") };
                    remove = new RadioButton { Content = "Rimuovi i commenti", GroupName = "cm", Margin = new Thickness(26, 2, 0, 2), Foreground = (Brush)FindResource("FgBrush") };
                    AutomationProperties.SetName(anon, "Anonimizza gli autori"); AutomationProperties.SetName(remove, "Rimuovi i commenti");
                    panel.Children.Add(anon); panel.Children.Add(remove);
                }
            }
        }
        panel.Children.Add(H("Non so pulire (restano nel file)"));
        foreach (var w in rep.Warnings)
            panel.Children.Add(new TextBlock { Text = "• " + w, TextWrapping = TextWrapping.Wrap, FontSize = 12, Margin = new Thickness(0, 1, 0, 1), Foreground = (Brush)FindResource("MutedBrush") });
        var scroll = new ScrollViewer { Content = panel, Width = 520, MaxHeight = 330, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Margin = new Thickness(38, 0, 0, 0) };
        var d = new ThemedDialog(this, "Scudo privacy", "Scegli cosa togliere. Si crea una copia: il file aperto non viene modificato. Tutto avviene in locale.", DialogKind.Info, scroll, mutedMessage: true);
        d.AddButton("Annulla", false, true);
        if (rep.Findings.Count > 0) d.AddButton("Crea copia pulita…", true);
        choice = () => (boxes.Where(b => b.Box.IsChecked == true).Select(b => b.Id).ToHashSet(), remove?.IsChecked == true ? CommentMode.Remove : CommentMode.Anonymize);
        return d;
    }

    void OnPrivacy()
    {
        if (_doc == null) return;
        CommitEdit(); SaveNotes();
        var data = _doc.Snapshot();   // stato corrente, anche se non salvato
        var rep = PrivacyShield.Scan(data);
        var d = BuildPrivacyDialog(rep, out var choice);
        d.ShowDialog();
        if (d.Result != 1 || rep.Findings.Count == 0) return;
        var (ids, mode) = choice();
        if (ids.Count == 0) { ThemedDialog.Notify(this, "Scudo privacy", "Non hai selezionato nulla da togliere: nessuna copia creata."); return; }
        var folder = _doc.FilePath != null ? Path.GetDirectoryName(_doc.FilePath) : null;
        var dlg = new SaveFileDialog { Filter = "Presentazione (*.pptx)|*.pptx", FileName = PrivacyShield.ProposeName(_doc.FilePath), InitialDirectory = folder ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), DefaultExt = ".pptx", Title = "Salva la copia ripulita" };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            var (clean, removed) = PrivacyShield.Clean(data, ids, mode);
            PrivacyShield.WriteCopy(clean, dlg.FileName, _doc.FilePath);
            var after = PrivacyShield.Scan(File.ReadAllBytes(dlg.FileName));   // ricontrollo della copia scritta
            var msg = new System.Text.StringBuilder();
            msg.AppendLine("Copia creata: " + dlg.FileName).AppendLine().AppendLine("Tolto:");
            foreach (var r in removed) msg.AppendLine("• " + r);
            msg.AppendLine().AppendLine(after.Findings.Count == 0 ? "Ricontrollo della copia: non trovo più i dati che so cercare." : "Ricontrollo della copia: ancora presenti");
            foreach (var f in after.Findings) msg.AppendLine($"• {f.Label}: {f.Value}");
            msg.AppendLine().AppendLine("Resta da verificare a mano:");
            foreach (var w in after.Warnings) msg.AppendLine("• " + w);
            var res = new ThemedDialog(this, "Scudo privacy", "Il file aperto non è stato modificato.", DialogKind.Info, new ScrollViewer { Width = 520, MaxHeight = 320, Margin = new Thickness(38, 8, 0, 0), VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = new TextBlock { Text = msg.ToString().TrimEnd(), TextWrapping = TextWrapping.Wrap, FontSize = 12.5, Foreground = (Brush)FindResource("FgBrush") } }, mutedMessage: true);
            res.AddButton("OK", true, true); res.ShowDialog();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
        { ThemedDialog.Notify(this, "Scudo privacy", "Copia non creata: " + ex.Message, DialogKind.Error); }
    }

    /// <summary>Per gli screenshot: mostra la finestra di analisi senza interazione.</summary>
    ThemedDialog ShowPrivacyForShot()
    {
        var rep = PrivacyShield.Scan(_doc!.Snapshot());
        var d = BuildPrivacyDialog(rep, out _); d.Show(); return d;
    }
}
