using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using Microsoft.Win32;
using PowerPresentations.Core;

namespace PowerPresentations.App;

public enum DialogKind { Info, Warning, Error, Question }

/// <summary>Finestra di dialogo propria: senza bordo di sistema, angoli arrotondati, colori del tema attivo e accento dell'app. Sostituisce MessageBox.</summary>
public sealed class ThemedDialog : Window
{
    readonly StackPanel _buttons = new() { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 22, 0, 0) };
    public int Result { get; private set; } = -1;
    Button? _primary; int _cancelIndex = -1;

    [System.Runtime.InteropServices.DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

    static Brush R(string k) => (Brush)Application.Current.Resources[k];

    public ThemedDialog(Window? owner, string title, string message, DialogKind kind, UIElement? extra = null, bool mutedMessage = false)
    {
        Owner = owner is { IsVisible: true } ? owner : null;
        WindowStyle = WindowStyle.None; AllowsTransparency = false; Background = R("PanelBrush");   // niente finestra a strati: con AllowsTransparency il testo è in software e sgranato
        System.Windows.Shell.WindowChrome.SetWindowChrome(this, new System.Windows.Shell.WindowChrome { CaptionHeight = 0, ResizeBorderThickness = new Thickness(0), GlassFrameThickness = new Thickness(0), CornerRadius = new CornerRadius(0), UseAeroCaptionButtons = false });
        SourceInitialized += (_, _) => { int pref = 2; DwmSetWindowAttribute(new System.Windows.Interop.WindowInteropHelper(this).Handle, 33, ref pref, 4); };   // angoli arrotondati e ombra dal sistema (Windows 11) ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;
        SizeToContent = SizeToContent.WidthAndHeight; WindowStartupLocation = Owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
        MaxWidth = 560; Title = title; AutomationProperties.SetName(this, title);
        var icon = Icons.Make(kind switch { DialogKind.Error => "close", DialogKind.Warning => "hide", DialogKind.Question => "notes", _ => "dir-r" }, 26,
            kind is DialogKind.Error or DialogKind.Warning ? R("PowersBrush") : R("AccentBrush"));
        var body = new StackPanel { Margin = new Thickness(24, 20, 24, 20), MinWidth = 440 };
        var head = new StackPanel { Orientation = Orientation.Horizontal };
        head.Children.Add(icon);
        head.Children.Add(new TextBlock { Text = title, FontSize = 17, FontWeight = FontWeights.SemiBold, Margin = new Thickness(12, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = R("FgBrush") });
        body.Children.Add(head);
        var msg = new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap, FontSize = mutedMessage ? 12.5 : 13.5, Margin = new Thickness(38, 6, 0, 0), Foreground = R(mutedMessage ? "MutedBrush" : "FgBrush") };
        AutomationProperties.SetName(msg, message); body.Children.Add(msg);
        if (extra != null) body.Children.Add(extra);
        body.Children.Add(_buttons);
        var card = new Border { Background = R("PanelBrush"), BorderBrush = R("LineBrush"), BorderThickness = new Thickness(1), Child = body };
        UseLayoutRounding = true; SnapsToDevicePixels = true;
        Content = card;
        MouseLeftButtonDown += (_, e) => { if (e.OriginalSource is not TextBox and not Button) { try { DragMove(); } catch (InvalidOperationException) { } } };
        KeyDown += (_, e) => { if (e.Key == Key.Escape && _cancelIndex >= 0) { Result = _cancelIndex; Close(); e.Handled = true; } };
        ContentRendered += (_, _) => { var w = ActualWidth; var h = ActualHeight; SizeToContent = SizeToContent.Manual; Width = w; Height = h; };   // fissa la misura: senza, con WindowChrome resta una fascia non dipinta a destra/sotto
        Loaded += (_, _) => { (_primary ?? _buttons.Children.OfType<Button>().FirstOrDefault())?.Focus(); };
    }

    /// <param name="validate">Chiamata prima di chiudere: ritorna false per lasciare aperto il dialogo (errore di validazione).</param>
    public Button AddButton(string label, bool primary = false, bool cancel = false, Func<bool>? validate = null)
    {
        int index = _buttons.Children.Count;
        var b = new Button
        {
            Content = label, MinWidth = 92, Height = 32, Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(14, 0, 14, 0),
            Foreground = primary ? Brushes.White : R("FgBrush"), Background = primary ? R("AccentBrush") : Brushes.Transparent, BorderBrush = primary ? R("AccentBrush") : R("LineBrush"), IsDefault = primary
        };
        AutomationProperties.SetName(b, label);
        b.Click += (_, _) => { if (validate != null && !validate()) return; Result = index; Close(); };
        if (primary) _primary = b;
        if (cancel) _cancelIndex = index;
        _buttons.Children.Add(b); return b;
    }

    /// <summary>Messaggio semplice (sostituisce MessageBox con OK).</summary>
    public static void Notify(Window? owner, string title, string message, DialogKind kind = DialogKind.Info)
    {
        var d = new ThemedDialog(owner, title, message, kind); d.AddButton("OK", true, true); d.ShowDialog();
    }

    /// <summary>Domanda con due scelte; ritorna true per la conferma.</summary>
    public static bool Confirm(Window? owner, string title, string message, string yes, string no = "Annulla")
    {
        var d = new ThemedDialog(owner, title, message, DialogKind.Question); d.AddButton(yes, true); d.AddButton(no, false, true); d.ShowDialog(); return d.Result == 0;
    }

    static string DefaultFolder() { if (Environment.GetEnvironmentVariable("PP_DEFAULT_FOLDER") is { Length: > 0 } o) return o;   // solo per screenshot senza dati personali
        var f = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); return Directory.Exists(f) ? f : Environment.CurrentDirectory; }

    /// <summary>Richiesta di salvataggio. Con Save ritorna il percorso scelto: qui non viene scritto nulla.</summary>
    public static ThemedDialog BuildSave(Window? owner, string? existingPath, bool saveAsMode, string suggestedName, out Func<(SaveChoice, string?)> result)
    {
        bool needsPath = existingPath == null || saveAsMode;
        string display = existingPath != null ? Path.GetFileName(existingPath) : suggestedName + SavePlan.Extension;
        TextBox? name = null, folder = null; TextBlock? err = null; ThemedDialog? d = null;
        UIElement? extra = null;
        if (needsPath)
        {
            var g = new StackPanel { Margin = new Thickness(0, 16, 0, 0) };
            TextBlock L(string t) => new() { Text = t, FontSize = 12, Foreground = R("MutedBrush"), Margin = new Thickness(0, 8, 0, 4) };
            name = new TextBox { Text = Path.GetFileNameWithoutExtension(display), FontSize = 13.5, Padding = new Thickness(6, 5, 6, 5) }; AutomationProperties.SetName(name, "Nome file");
            folder = new TextBox { Text = existingPath != null ? Path.GetDirectoryName(existingPath) ?? DefaultFolder() : DefaultFolder(), FontSize = 12.5, Padding = new Thickness(6, 5, 6, 5) }; AutomationProperties.SetName(folder, "Percorso");
            string browseLabel = saveAsMode ? "Altre opzioni…" : "Sfoglia…";
            var browse = new Button { Content = browseLabel, Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(14, 0, 14, 0), MinWidth = 96 }; AutomationProperties.SetName(browse, browseLabel);
            var row = new DockPanel(); DockPanel.SetDock(browse, Dock.Right); row.Children.Add(browse); row.Children.Add(folder);
            err = new TextBlock { Foreground = R("PowersBrush"), FontSize = 12, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 8, 0, 0), Visibility = Visibility.Collapsed }; AutomationProperties.SetName(err, "Errore");
            g.Children.Add(L("Nome file")); g.Children.Add(name); g.Children.Add(L("Percorso")); g.Children.Add(row); g.Children.Add(err);
            browse.Click += (_, _) =>
            {
                var sd = new SaveFileDialog { Filter = "Presentazione (*.pptx)|*.pptx", FileName = name.Text, InitialDirectory = Directory.Exists(folder.Text) ? folder.Text : DefaultFolder(), DefaultExt = ".pptx", OverwritePrompt = false };
                if (sd.ShowDialog(d) == true) { name.Text = Path.GetFileNameWithoutExtension(sd.FileName); folder.Text = Path.GetDirectoryName(sd.FileName) ?? folder.Text; }
            };
            extra = g;
        }
        string title = saveAsMode ? "Salva con nome" : $"Vuoi salvare le modifiche a «{display}»?";
        string msg = saveAsMode ? "Scegli nome e cartella della presentazione." : "Se non salvi, le modifiche andranno perse.";
        d = new ThemedDialog(owner, title, msg, DialogKind.Question, extra, mutedMessage: true);
        string? chosen = null;
        bool Validate()
        {
            if (!needsPath) { chosen = existingPath; return true; }
            var (e, overwrite, p) = SavePlan.Check(folder!.Text.Trim(), name!.Text);
            if (e != null) { err!.Text = e; err.Visibility = Visibility.Visible; return false; }
            bool same = existingPath != null && string.Equals(Path.GetFullPath(existingPath), Path.GetFullPath(p), StringComparison.OrdinalIgnoreCase);
            if (overwrite && !same && !Confirm(d, "Sostituire il file?", $"«{Path.GetFileName(p)}» esiste già in questa cartella. Vuoi sostituirlo?", "Sostituisci")) return false;
            chosen = p; return true;
        }
        // ordine come in Windows 11: Annulla, Non salvare, Salva (primario a destra)
        d.AddButton("Annulla", false, true);
        if (!saveAsMode) d.AddButton("Non salvare");
        d.AddButton("Salva", true, false, Validate);
        d.Loaded += (_, _) => { if (name != null) { name.Focus(); name.SelectAll(); } };
        var dlg = d;
        result = () =>
        {
            int dont = saveAsMode ? -1 : 1, save = saveAsMode ? 1 : 2;
            var choice = dlg.Result == save ? SaveChoice.Save : dlg.Result == dont ? SaveChoice.DontSave : SaveChoice.Cancel;
            return (choice, choice == SaveChoice.Save ? chosen : null);
        };
        return d;
    }

    public static (SaveChoice Choice, string? Path) AskSave(Window? owner, string? existingPath, bool saveAsMode, string suggestedName = "Presentazione")
    {
        var d = BuildSave(owner, existingPath, saveAsMode, suggestedName, out var result); d.ShowDialog(); return result();
    }
}
