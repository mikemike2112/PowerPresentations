using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.App;

/// <summary>Comandi del ribbon aggiunti dopo la prima versione: appunti, selezione, menu contestuali.</summary>
public partial class MainWindow
{
    const string ClipFmt = "PowerPresentations.Shapes";
    (uint Uid, bool Cut)? _slideClip;

    /// <summary>Esegue una modifica con voce di annullamento; se fallisce lo stato torna com'era e non resta nessuna voce vuota.</summary>
    T Guard<T>(Func<T> op)
    {
        var before = _doc!.Snapshot();
        try { var r = op(); _undo.Add(before); if (_undo.Count > 40) _undo.RemoveAt(0); _redo.Clear(); UpdateQuick(); return r; }
        catch { _doc.Restore(before); _r = new SlideRenderer(_doc); throw; }
    }

    void Guard(Action op) => Guard<int>(() => { op(); return 0; });

    /// <summary>Come Guard ma mostra l'errore all'utente invece di propagarlo. Ritorna false se la modifica è fallita.</summary>
    bool Try(Action op, string title)
    {
        try { Guard(op); return true; }
        catch (Exception ex) when (ex is PptxFormatException or KeyNotFoundException or NotSupportedException or InvalidOperationException or ArgumentException or IOException)
        { ThemedDialog.Notify(this, title, ex.Message, DialogKind.Warning); RefreshCurrent(); return false; }
    }

    // ---------- appunti ----------
    static T? Retry<T>(Func<T> f)
    {
        for (int i = 0; i < 8; i++) { try { return f(); } catch (ExternalException) { Thread.Sleep(40); } }
        return default;
    }

    bool CopySelection()
    {
        var ids = View.SelectedIds;
        if (_doc == null || ids.Count == 0) { Status.Text = "Seleziona prima una o più forme."; return false; }
        try
        {
            var clip = _doc.CopyShapes(_cur, ids);
            var data = new DataObject(); data.SetData(ClipFmt, clip.ToBytes()); if (clip.Text.Length > 0) data.SetText(clip.Text);
            if (Retry(() => { Clipboard.SetDataObject(data, true); return true; }) != true) { Status.Text = "Appunti non disponibili: riprova."; return false; }
            Status.Text = ids.Count == 1 ? "Forma copiata." : $"{ids.Count} forme copiate."; return true;
        }
        catch (Exception ex) when (ex is PptxFormatException or KeyNotFoundException) { ThemedDialog.Notify(this, "Copia non riuscita", ex.Message, DialogKind.Warning); return false; }
    }

    void OnCopy() { if (SlidePanelFocused()) CopySlide(false); else CopySelection(); }

    void OnCut()
    {
        if (SlidePanelFocused()) { CopySlide(true); return; }
        if (!CopySelection()) return;
        OnDeleteShape(this, new RoutedEventArgs());
    }

    void OnPaste()
    {
        if (_doc == null) return;
        if (SlidePanelFocused() && _slideClip != null) { PasteSlide(); return; }
        var d = Retry(Clipboard.GetDataObject);
        if (d == null) { Status.Text = "Appunti non disponibili: riprova."; return; }
        if (d.GetDataPresent(ClipFmt) && d.GetData(ClipFmt) is byte[] b && ShapeClip.FromBytes(b) is { } clip)
        {
            List<int> ids = new();
            if (Try(() => ids = _doc.PasteShapes(_cur, clip), "Incolla non riuscito")) { RefreshCurrent(); View.SelectMany(ids); UpdateStatus(); View.Focus(); }
        }
        else if (d.GetDataPresent(DataFormats.Bitmap) && Retry(Clipboard.GetImage) is { } img)
        {
            var enc = new PngBitmapEncoder(); enc.Frames.Add(BitmapFrame.Create(img)); using var ms = new MemoryStream(); enc.Save(ms);
            try { AddPictureFit(ms.ToArray(), "Immagine incollata"); } catch (PptxFormatException ex) { ThemedDialog.Notify(this, "Incolla non riuscito", ex.Message, DialogKind.Warning); }
        }
        else if (d.GetDataPresent(DataFormats.UnicodeText) && d.GetData(DataFormats.UnicodeText) is string t && t.Trim().Length > 0)
        {
            t = t.Length > 20000 ? t[..20000] : t;   // ponytail: tetto al testo incollato → paginare
            var (w, h) = _doc.SlideSize; int lines = t.Split('\n').Length;
            int id = 0; Guard(() => id = _doc.AddTextBox(_cur, w / 2 - 200, Math.Max(10, h / 2 - 12 * lines), 400, 24 * lines + 8, t.TrimEnd('\r', '\n')));
            RefreshCurrent(); View.SelectedId = id; View.Focus();
        }
        else Status.Text = "Negli appunti non c'è nulla da incollare.";
    }

    bool CanPaste()
    {
        var d = Retry(Clipboard.GetDataObject); if (d == null) return false;
        return (SlidePanelFocused() && _slideClip != null) || d.GetDataPresent(ClipFmt) || d.GetDataPresent(DataFormats.Bitmap) || d.GetDataPresent(DataFormats.UnicodeText);
    }

    void ClipboardKey(Key k)
    {
        if (k == Key.C) OnCopy(); else if (k == Key.X) OnCut(); else OnPaste();
    }

    bool SlidePanelFocused() => Thumbs.IsKeyboardFocusWithin && !View.IsKeyboardFocusWithin;

    void DuplicateSelection()
    {
        var ids = View.SelectedIds.ToList(); if (_doc == null || ids.Count == 0) return;
        List<int> n = new();
        if (Try(() => n = _doc.PasteShapes(_cur, _doc.CopyShapes(_cur, ids)), "Duplica non riuscito")) { RefreshCurrent(); View.SelectMany(n); UpdateStatus(); }
    }

    // ---------- slide negli appunti (interni all'app) ----------
    void CopySlide(bool cut)
    {
        if (_doc == null) return;
        if (cut && _doc.Slides.Count <= 1) { Status.Text = "Non si può tagliare l'ultima slide."; return; }
        _slideClip = (_doc.SlideUid(_cur), cut); Status.Text = cut ? "Slide tagliata: incolla per spostarla." : "Slide copiata.";
    }

    void PasteSlide()
    {
        if (_doc == null || _slideClip is not { } c) return;
        int from = _doc.IndexOfUid(c.Uid); if (from < 0) { Status.Text = "La slide copiata non esiste più."; _slideClip = null; return; }
        SaveNotes(); CommitEdit();
        uint curUid = _doc.SlideUid(_cur); int newIdx;
        Snap();
        if (c.Cut)
        {
            int cur = _cur; if (from == cur) return;
            _doc.MoveSlide(from, from < cur ? cur : cur + 1); newIdx = from < cur ? cur : cur + 1; _slideClip = null;
        }
        else
        {
            int n = _doc.DuplicateSlide(from); int cn = _doc.IndexOfUid(curUid);
            if (n > cn) { _doc.MoveSlide(n, cn + 1); newIdx = cn + 1; } else { _doc.MoveSlide(n, cn); newIdx = cn; }
        }
        _cur = newIdx; RebuildThumbs(); ShowSlide(newIdx); UpdateTitle();
    }

    // ---------- selezione ----------
    void SelectAllShapes()
    {
        if (_doc == null) return;
        var ids = _doc.ShapeIds(_cur); if (ids.Count == 0) { Status.Text = "La slide non ha forme."; return; }
        View.SelectMany(ids); UpdateStatus(); View.Focus();
    }

    /// <summary>Seleziona: tutto, oppure una forma scelta dall'elenco (riquadro di selezione).</summary>
    void OnSelectMenu()
    {
        if (_doc == null) return;
        var ids = _doc.ShapeIds(_cur);
        var labels = new List<string> { "Seleziona tutto" }; labels.AddRange(ids.Select(i => _doc.ShapeLabel(_cur, i)));
        int k = ListPicker.Pick(this, "Seleziona", labels); if (k < 0) return;
        if (k == 0) SelectAllShapes(); else { View.SelectedId = ids[k - 1]; OnSelect(ids[k - 1]); View.Focus(); }
    }

    // ---------- menu contestuali ----------
    void BuildContextMenus()
    {
        MenuItem Item(string h, Action a) { var m = new MenuItem { Header = h }; AutomationProperties_Name(m, h); m.Click += (_, _) => a(); return m; }
        var shape = new ContextMenu();
        var cut = Item("Taglia", OnCut); var copy = Item("Copia", OnCopy); var paste = Item("Incolla", OnPaste); var dup = Item("Duplica", DuplicateSelection); var del = Item("Elimina", () => OnDeleteShape(this, new RoutedEventArgs()));
        foreach (var m in new[] { cut, copy, paste, dup, del }) shape.Items.Add(m);
        shape.Opened += (_, _) => { bool any = View.SelectedIds.Count > 0; cut.IsEnabled = copy.IsEnabled = dup.IsEnabled = del.IsEnabled = any; paste.IsEnabled = CanPaste(); };
        View.ContextMenu = shape;

        var thumb = new ContextMenu();
        var tn = Item("Nuova diapositiva", () => OnNewSlide(this, new RoutedEventArgs())); var td = Item("Duplica diapositiva", () => OnDuplicate(this, new RoutedEventArgs()));
        var tx = Item("Taglia", () => CopySlide(true)); var tc = Item("Copia", () => CopySlide(false)); var tp = Item("Incolla", PasteSlide);
        var th = Item("Nascondi/mostra", () => OnToggleHidden(this, new RoutedEventArgs())); var te = Item("Elimina diapositiva", () => OnDeleteSlide(this, new RoutedEventArgs()));
        foreach (var m in new[] { tn, td, tx, tc, tp, th, te }) thumb.Items.Add(m);
        thumb.Opened += (_, _) => tp.IsEnabled = _slideClip != null;
        Thumbs.ContextMenu = thumb;
        Thumbs.PreviewMouseRightButtonDown += (_, e) =>
        {
            if (e.OriginalSource is DependencyObject o && ItemsControl.ContainerFromElement(Thumbs, o) is ListBoxItem c) { Thumbs.SelectedIndex = Thumbs.ItemContainerGenerator.IndexFromContainer(c); Thumbs.Focus(); }
        };
    }

    static void AutomationProperties_Name(DependencyObject d, string name) => System.Windows.Automation.AutomationProperties.SetName(d, name);

    // ---------- elenchi ----------
    /// <summary>Elenco puntato/numerato: se tutti i paragrafi lo hanno già, lo toglie (ponytail: un elenco ereditato dal layout conta come "non impostato" → un clic lo rende esplicito).</summary>
    void Bullet(string kind)
    {
        int id = View.SelectedId;
        if (_doc == null || id == 0 || _doc.TextBodies(_cur, id).Count == 0) { Status.Text = "Seleziona prima una forma con testo."; return; }
        var target = _doc.GetBullet(_cur, id) == kind ? "none" : kind;
        Snap(); _doc.SetBullet(_cur, id, target); RefreshCurrent(); View.SelectedId = id;
    }

    // ---------- transizioni ----------
    void OnApplyTransAll()
    {
        if (_doc == null) return;
        if (_doc.Slides.Count < 2) { Status.Text = "C'è una sola slide."; return; }
        Snap(); _doc.ApplyTransitionToAll(_cur); RebuildThumbs(); UpdateTitle(); Status.Text = "Transizione applicata a tutte le slide.";
    }
}
