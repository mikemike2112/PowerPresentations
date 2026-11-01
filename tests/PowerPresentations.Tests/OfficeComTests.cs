using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Tests;

/// <summary>Esegue tools/office-inspect.ps1: apre i file nel VERO PowerPoint via COM (WithWindow=false, watchdog per PID) e ritorna il JSON.</summary>
public static class ComRunner
{
    public static Dictionary<string, JsonElement> Inspect(IEnumerable<string> files, string? pdfDir = null, string? pngDir = null)
    {
        var list = Path.Combine(Env.Work, "com-list-" + Guid.NewGuid().ToString("N") + ".txt");
        var json = Path.ChangeExtension(list, ".json");
        File.WriteAllLines(list, files.Select(Path.GetFullPath), new UTF8Encoding(true));
        var args = $"-NoProfile -ExecutionPolicy Bypass -File \"{Path.Combine(Env.Tools, "office-inspect.ps1")}\" -ListFile \"{list}\" -OutJson \"{json}\" -TimeoutSec 25"
            + (pdfDir != null ? $" -PdfDir \"{pdfDir}\"" : "") + (pngDir != null ? $" -PngDir \"{pngDir}\"" : "");
        using var p = Process.Start(new ProcessStartInfo("powershell", args) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true })!;
        var so = p.StandardOutput.ReadToEndAsync(); var se = p.StandardError.ReadToEndAsync();
        if (!p.WaitForExit(20 * 60 * 1000)) { try { p.Kill(true); } catch (InvalidOperationException) { } throw new TimeoutException("office-inspect.ps1 non è terminato."); }
        if (!File.Exists(json)) throw new InvalidOperationException("Nessun JSON prodotto: " + so.Result + se.Result);
        using var doc = JsonDocument.Parse(File.ReadAllBytes(json));
        return doc.RootElement.EnumerateArray().ToDictionary(e => Path.GetFullPath(e.GetProperty("path").GetString()!), e => e.Clone(), StringComparer.OrdinalIgnoreCase);
    }

    public static bool Ok(JsonElement r) => r.GetProperty("ok").GetBoolean();
    public static string Error(JsonElement r) => r.GetProperty("error").GetString() ?? "";
    public static JsonElement Content(JsonElement r) => r.GetProperty("content");

    /// <summary>Forma testuale del contenuto letto da Office (testo, valori, n. slide e forme) per il confronto.</summary>
    public static string Signature(JsonElement r)
    {
        var c = Content(r); var sb = new StringBuilder();
        sb.Append($"slides={c.GetProperty("slideCount")};{c.GetProperty("width")}x{c.GetProperty("height")};sections={c.GetProperty("sections")}\n");
        foreach (var s in c.GetProperty("slides").EnumerateArray())
        {
            sb.Append($"S{s.GetProperty("index")} layout={s.GetProperty("layout")} hidden={s.GetProperty("hidden")} tr={s.GetProperty("transition")} notes=[{s.GetProperty("notes")}] comments={s.GetProperty("comments")} shapes={s.GetProperty("shapes").GetArrayLength()}\n");
            foreach (var sh in s.GetProperty("shapes").EnumerateArray())
            {
                sb.Append($"  {sh.GetProperty("name")}|{sh.GetProperty("type")}|[{sh.GetProperty("text")}]|{sh.GetProperty("l")},{sh.GetProperty("t")},{sh.GetProperty("w")},{sh.GetProperty("h")}");
                if (sh.TryGetProperty("table", out var t) && t.ValueKind == JsonValueKind.Object) sb.Append($"|T{t.GetProperty("rows")}x{t.GetProperty("cols")}:{string.Join(",", t.GetProperty("cells").EnumerateArray().Select(x => x.GetString()))}");
                if (sh.TryGetProperty("group", out var g)) sb.Append($"|G{g}");
                sb.Append('\n');
            }
        }
        return sb.ToString();
    }

    public static JsonElement Slide(JsonElement r, int i) => Content(r).GetProperty("slides")[i];
}

public sealed class ComFixture
{
    public Dictionary<string, JsonElement> Originals { get; }
    public Dictionary<string, JsonElement> Saved { get; }
    public string SavedDir { get; }

    public ComFixture()
    {
        SavedDir = Directory.CreateDirectory(Path.Combine(Env.Work, "saved")).FullName;
        var files = Directory.GetFiles(Env.Corpus, "*.pptx").OrderBy(f => f).ToList();
        foreach (var f in files) { using var d = PptxDocument.Open(f); d.SaveAs(Path.Combine(SavedDir, Path.GetFileName(f))); }
        Originals = ComRunner.Inspect(files);
        Saved = ComRunner.Inspect(files.Select(f => Path.Combine(SavedDir, Path.GetFileName(f))));
    }
}

[CollectionDefinition("Desktop", DisableParallelization = true)]
public class OfficeComCollection { }

[Collection("Desktop"), Trait("Category", "OfficeCom")]
public class OfficeRoundTripTests : IClassFixture<ComFixture>
{
    readonly ComFixture _f;
    public OfficeRoundTripTests(ComFixture f) { _f = f; }

    [Theory, MemberData(nameof(Env.CorpusFiles), MemberType = typeof(Env))]
    public void Original_OpensInOffice(string name)
    {
        var r = _f.Originals[Path.GetFullPath(Env.CorpusPath(name))];
        Assert.True(ComRunner.Ok(r), $"{name}: l'originale non si apre in Office: {ComRunner.Error(r)}");
    }

    // b. il salvato si riapre in Office senza errori né ripristino; c. contenuto identico letto da Office
    [Theory, MemberData(nameof(Env.CorpusFiles), MemberType = typeof(Env))]
    public void SavedByClone_OpensInOfficeWithIdenticalContent(string name)
    {
        var r = _f.Saved[Path.GetFullPath(Path.Combine(_f.SavedDir, name))];
        Assert.True(ComRunner.Ok(r), $"{name}: Office non apre il file salvato: {ComRunner.Error(r)}");
        var o = _f.Originals[Path.GetFullPath(Env.CorpusPath(name))];
        Assert.Equal(ComRunner.Signature(o), ComRunner.Signature(r));
    }
}

/// <summary>e. Controllo negativo: il rilevatore DEVE fallire su file volutamente corrotti; senza questo il round-trip non vale.</summary>
[Collection("Desktop"), Trait("Category", "OfficeCom")]
public class OfficeNegativeControlTests
{
    static void Mutate(string src, string dst, Func<string, byte[], byte[]?> fn)
    {
        using var zi = ZipFile.OpenRead(src); using var zo = ZipFile.Open(dst, ZipArchiveMode.Create);
        foreach (var e in zi.Entries)
        {
            using var ms = new MemoryStream(); using (var s = e.Open()) s.CopyTo(ms);
            var r = fn(e.FullName, ms.ToArray()); if (r == null) continue;
            using var w = zo.CreateEntry(e.FullName).Open(); w.Write(r);
        }
    }

    [Fact]
    public void Detector_FailsOnBrokenXml_MissingPart_BadAttribute_AndPassesGoodFile()
    {
        var src = Env.CorpusPath("mixed-realistic-1.pptx"); var dir = Directory.CreateDirectory(Path.Combine(Env.Work, "neg")).FullName;
        var bad1 = Path.Combine(dir, "bad-xml.pptx"); var bad2 = Path.Combine(dir, "bad-rel.pptx"); var bad3 = Path.Combine(dir, "bad-attr.pptx"); var good = Path.Combine(dir, "good.pptx");
        File.Copy(src, good, true);
        Mutate(src, bad1, (n, d) => n == "ppt/slides/slide2.xml" ? Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(d).Replace("<p:cSld>", "<p:cSld><p:oops>")) : d);
        Mutate(src, bad2, (n, d) => n == "ppt/slideLayouts/slideLayout2.xml" ? null : d);
        Mutate(src, bad3, (n, d) => n == "ppt/slides/slide3.xml" ? Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(d).Replace("<a:off x=\"", "<a:off x=\"abc")) : d);
        var r = ComRunner.Inspect(new[] { good, bad1, bad2, bad3 });
        Assert.True(ComRunner.Ok(r[good]), "il file buono deve passare: " + ComRunner.Error(r[good]));
        foreach (var b in new[] { bad1, bad2, bad3 }) Assert.False(ComRunner.Ok(r[b]), $"{Path.GetFileName(b)}: il rilevatore NON ha colto la corruzione");
    }

    [Fact]
    public void Watchdog_KillsOnlyOwnedInstance_WhenHeartbeatStops()
    {
        var p = Process.Start(new ProcessStartInfo("powershell", $"-NoProfile -ExecutionPolicy Bypass -File \"{Path.Combine(Env.Tools, "test-watchdog.ps1")}\"") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true })!;
        var o = p.StandardOutput.ReadToEnd(); p.WaitForExit(60000);
        Assert.Contains(p.ExitCode == 2 ? "SKIP" : "WATCHDOG_OK", o); // 2 = PowerPoint dell'utente già aperto: non si tocca
    }
}

/// <summary>f. Test con MODIFICA: il clone modifica e salva, Office verifica che la modifica ci sia e che il resto non sia perso.</summary>
public sealed class ComEditFixture
{
    public Dictionary<string, JsonElement> R { get; }
    public string Dir { get; }
    public static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    public string P(string n) => Path.Combine(Dir, n + ".pptx");

    static int FirstText(PptxDocument d, int slide) => new SlideRenderer(d).Render(slide).Hits.First(h => h.HasText).Id;

    public ComEditFixture()
    {
        Dir = Directory.CreateDirectory(Path.Combine(Env.Work, "edit")).FullName;
        void Edit(string src, string name, Action<PptxDocument> a) { using var d = PptxDocument.Open(Env.CorpusPath(src)); a(d); d.SaveAs(P(name)); }
        Edit("textbox-basic.pptx", "text", d => d.SetShapeText(0, FirstText(d, 0), "Testo cambiato dal clone"));
        Edit("mixed-realistic-1.pptx", "reorder", d => d.MoveSlide(2, 0));
        Edit("title-slide.pptx", "addslide", d => { var i = d.AddSlide(d.Layouts[1], 0); d.SetShapeText(i, FirstText(d, i), "Nuova slide"); });
        Edit("sections.pptx", "delslide", d => d.DeleteSlide(1));
        Edit("textbox-basic.pptx", "notes", d => d.SetNotes(0, "Nota scritta dal clone"));
        Edit("title-slide.pptx", "transition", d => d.SetTransition(0, "push", "l", "fast"));
        Edit("title-slide.pptx", "objects", d =>
        {
            d.AddTextBox(0, 50, 400, 300, 40, "Casella nuova"); d.AddShape(0, "ellipse", 400, 380, 150, 80, "Forma");
            d.AddPicture(0, Png, 600, 380, 100, 80); var t = d.AddTable(0, 2, 3, 50, 100, 400); d.SetCellText(0, t, 0, 1, "Cella B1");
        });
        Edit("mixed-realistic-1.pptx", "duplicate", d => d.DuplicateSlide(1));
        Edit("hidden-slide.pptx", "unhide", d => d.SetHidden(1, false));
        Edit("notes-speaker.pptx", "editnotes", d => d.SetNotes(1, "Seconda nota cambiata"));
        Edit("table-merged-cells.pptx", "cell", d => { var gf = new SlideRenderer(d).Render(0).Hits.First(h => h.HasText); d.SetCellText(0, gf.Id, 3, 0, "modificata"); });
        using (var d = PptxDocument.New()) { d.SetShapeText(0, FirstText(d, 0), "Da modello"); d.SaveAs(P("new")); }
        R = ComRunner.Inspect(new[] { "text", "reorder", "addslide", "delslide", "notes", "transition", "objects", "duplicate", "unhide", "editnotes", "cell", "new" }.Select(P));
    }
}

[Collection("Desktop"), Trait("Category", "OfficeCom")]
public class OfficeEditTests : IClassFixture<ComEditFixture>
{
    readonly ComEditFixture _f;
    public OfficeEditTests(ComEditFixture f) { _f = f; }
    JsonElement Get(string n) { var r = _f.R[Path.GetFullPath(_f.P(n))]; Assert.True(ComRunner.Ok(r), $"{n}: Office non lo apre: {ComRunner.Error(r)}"); return r; }
    static string Txt(JsonElement r, int slide) => string.Join("|", ComRunner.Slide(r, slide).GetProperty("shapes").EnumerateArray().Select(s => s.GetProperty("text").GetString()));

    [Fact] public void TextEdit_IsVisibleInOffice_AndRestPreserved()
    {
        var r = Get("text"); Assert.Contains("Testo cambiato dal clone", Txt(r, 0)); Assert.Contains("Un testo molto più lungo", Txt(r, 0));
        Assert.Equal(2, ComRunner.Slide(r, 0).GetProperty("shapes").GetArrayLength());
    }
    [Fact] public void Reorder_PreservesNotes_AndOrder()
    {
        var r = Get("reorder"); Assert.Equal(3, ComRunner.Content(r).GetProperty("slideCount").GetInt32());
        Assert.Contains("Risultati", Txt(r, 0)); Assert.Contains("Piano trimestrale", Txt(r, 1)); Assert.Contains("Obiettivi", Txt(r, 2));
        Assert.Equal("Spiegare i tre obiettivi", ComRunner.Slide(r, 2).GetProperty("notes").GetString());
    }
    [Fact] public void AddSlide_AppearsWithLayoutAndText()
    {
        var r = Get("addslide"); Assert.Equal(2, ComRunner.Content(r).GetProperty("slideCount").GetInt32()); Assert.Contains("Nuova slide", Txt(r, 1));
        Assert.Contains("Titolo della presentazione", Txt(r, 0));
    }
    [Fact] public void DeleteSlide_WithSections_OfficeOpensClean() { var r = Get("delslide"); Assert.Equal(3, ComRunner.Content(r).GetProperty("slideCount").GetInt32()); Assert.DoesNotContain("Slide 2", string.Join("", Enumerable.Range(0, 3).Select(i => Txt(r, i)))); }
    [Fact] public void Notes_CreatedByClone_ReadByOffice() => Assert.Equal("Nota scritta dal clone", ComRunner.Slide(Get("notes"), 0).GetProperty("notes").GetString());
    [Fact] public void Notes_EditedByClone_ReadByOffice_OtherNotesIntact()
    {
        var r = Get("editnotes"); Assert.Equal("Seconda nota cambiata", ComRunner.Slide(r, 1).GetProperty("notes").GetString());
        Assert.Contains("Prima nota del relatore", ComRunner.Slide(r, 0).GetProperty("notes").GetString());
    }
    [Fact] public void Transition_WrittenByClone_ReadByOffice() => Assert.NotEqual(0, ComRunner.Slide(Get("transition"), 0).GetProperty("transition").GetInt32());
    [Fact] public void InsertedObjects_AreRealOfficeShapes()
    {
        var shapes = ComRunner.Slide(Get("objects"), 0).GetProperty("shapes").EnumerateArray().ToList();
        Assert.Contains(shapes, s => s.GetProperty("type").GetInt32() == 17 && s.GetProperty("text").GetString() == "Casella nuova");       // msoTextBox
        Assert.Contains(shapes, s => s.GetProperty("type").GetInt32() == 1 && s.GetProperty("text").GetString() == "Forma");                // msoAutoShape
        Assert.Contains(shapes, s => s.GetProperty("type").GetInt32() == 13);                                                               // msoPicture
        var t = shapes.First(s => s.GetProperty("type").GetInt32() == 19).GetProperty("table");                                             // msoTable
        Assert.Equal(2, t.GetProperty("rows").GetInt32()); Assert.Equal(3, t.GetProperty("cols").GetInt32()); Assert.Equal("Cella B1", t.GetProperty("cells")[1].GetString());
    }
    [Fact] public void DuplicateSlide_HasSameContent() { var r = Get("duplicate"); Assert.Equal(4, ComRunner.Content(r).GetProperty("slideCount").GetInt32()); Assert.Equal(Txt(r, 1), Txt(r, 2)); Assert.Equal("Spiegare i tre obiettivi", ComRunner.Slide(r, 2).GetProperty("notes").GetString()); }
    [Fact] public void Unhide_ReadByOffice() => Assert.Equal(0, ComRunner.Slide(Get("unhide"), 1).GetProperty("hidden").GetInt32());
    [Fact] public void TableCellEdit_Merged_ReadByOffice()
    {
        var shapes = ComRunner.Slide(Get("cell"), 0).GetProperty("shapes").EnumerateArray().ToList();
        var cells = shapes.First(s => s.GetProperty("type").GetInt32() == 19).GetProperty("table").GetProperty("cells").EnumerateArray().Select(x => x.GetString()).ToList();
        Assert.Equal("modificata", cells[12]);
    }
    [Fact] public void NewFromTemplate_OpensInOffice() { var r = Get("new"); Assert.Contains("Da modello", Txt(r, 0)); }
}
