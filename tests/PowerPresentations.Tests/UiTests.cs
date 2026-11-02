using System.Diagnostics;
using System.Text.Json;

namespace PowerPresentations.Tests;

/// <summary>Prove sulla UI reale (WPF) avviata come processo: UI Automation, input reale del mouse/tastiera, screenshot. Solo processi PowerPresentations.App avviati dal test.</summary>
/// <summary>Blocco condiviso con le altre app: chi ha ui.lock può usare mouse e tastiera reali. Una cartella più vecchia di 20 minuti è abbandonata.</summary>
public sealed class UiLockFixture : IDisposable
{
    const string Lock = @"C:\dev\_office-ref\ui.lock";
    readonly bool _own;
    public UiLockFixture()
    {
        var deadline = DateTime.UtcNow.AddMinutes(45);
        while (true)
        {
            // controllo+creazione atomici fra processi (anche altre app della suite): Mutex di sistema tenuto solo qui
            using (var m = new Mutex(false, @"Global\PowerSuiteUiLock"))
            {
                bool got; try { got = m.WaitOne(TimeSpan.FromSeconds(30)); } catch (AbandonedMutexException) { got = true; }
                if (got)
                {
                    try
                    {
                        try { if (Directory.Exists(Lock) && DateTime.UtcNow - Directory.GetCreationTimeUtc(Lock) > TimeSpan.FromMinutes(20)) Directory.Delete(Lock); } catch (IOException) { }   // blocco abbandonato
                        try { if (!Directory.Exists(Lock)) { Directory.CreateDirectory(Lock); _own = true; return; } } catch (IOException) { }
                    }
                    finally { m.ReleaseMutex(); }
                }
            }
            if (DateTime.UtcNow > deadline) throw new TimeoutException("Blocco UI non ottenuto: un'altra app sta usando mouse e tastiera.");
            Thread.Sleep(5000);
        }
    }
    public void Dispose() { if (_own) { try { Directory.Delete(Lock); } catch (IOException) { } } }
}

[Collection("Desktop"), Trait("Category", "Ui")]
public class UiTests : IClassFixture<UiLockFixture>
{
    public UiTests(UiLockFixture _) { }

    static string Run(string script, string args, int timeoutMs = 120000)
    {
        var exe = Path.Combine(Env.Repo, "src", "PowerPresentations.App", "bin", "Debug", "net10.0-windows", "PowerPresentations.App.exe");
        Assert.True(File.Exists(exe), "Compilare prima l'app: dotnet build");
        using var p = Process.Start(new ProcessStartInfo("powershell", $"-NoProfile -ExecutionPolicy Bypass -File \"{Path.Combine(Env.Tools, script)}\" {args}") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true })!;
        var so = p.StandardOutput.ReadToEndAsync(); var se = p.StandardError.ReadToEndAsync();
        if (!p.WaitForExit(timeoutMs)) { try { p.Kill(true); } catch (InvalidOperationException) { } Assert.Fail("script UI non terminato: " + script); }
        Assert.True(p.ExitCode == 0, se.Result + so.Result);
        return so.Result;
    }

    [Fact]
    public void RealWindow_Transition_Duplicate_DragReorder_Save_Fullscreen_Show()
    {
        var shots = Directory.CreateDirectory(Path.Combine(Env.Work, "ui")).FullName;
        var json = Run("ui-smoke.ps1", $"-Shots \"{shots}\"").Trim().Split('\n').Last().Trim();
        using var d = JsonDocument.Parse(json); var r = d.RootElement;
        Assert.Equal(4, r.GetProperty("thumbnails").GetInt32());                                  // duplica
        Assert.Equal(4, r.GetProperty("savedSlideCount").GetInt32());                             // salvato con Ctrl+S
        Assert.Equal("Piano trimestrale | Obiettivi | Piano trimestrale | Risultati", r.GetProperty("savedTitles").GetString()); // drag & drop reale: copia davanti, originale spostata in 3ª
        Assert.Equal("True,False,True,True", r.GetProperty("savedFade").GetString());             // transizione applicata dal clic e copiata dalla duplicazione
        Assert.True(r.GetProperty("showWindow").GetBoolean());                                    // F5: finestra "Presentazione"
        Assert.Matches(@"^\d{3,}x\d{3,}$", r.GetProperty("showSize").GetString()!); // dimensione della finestra di presentazione (a schermo intero: pari allo schermo principale)
        Assert.True(r.GetProperty("closed").GetBoolean());                                        // chiusura senza dialoghi: documento non sporco dopo il salvataggio
        Assert.True(new FileInfo(Path.Combine(shots, "app-schermo-intero-reale-2.png")).Length > 5000);
    }

    [Fact]
    public void PowersTab_RealClickTogglesTheme_PersistsAcrossRestart_CorruptSettingsDoNotCrash()
    {
        var shots = Directory.CreateDirectory(Path.Combine(Env.Work, "ui-powers")).FullName;
        var json = Run("ui-powers.ps1", $"-Shots \"{shots}\"").Trim().Split((char)10).Last().Trim();
        using var d = JsonDocument.Parse(json); var r = d.RootElement;
        Assert.True(r.GetProperty("defaultDarkInFile").GetBoolean());                          // prima apertura: nessun file, tema scuro
        Assert.Contains("\"Dark\":false", r.GetProperty("settingsAfterToggle").GetString());   // clic reale sull'interruttore -> salvato
        Assert.True(r.GetProperty("restartLuminance").GetInt32() > 180);                       // riavvio: parte chiaro
        Assert.Equal(0, r.GetProperty("corruptExit").GetInt32());                              // file corrotto: nessun crash
        Assert.True(r.GetProperty("corruptLuminance").GetInt32() < 60);                        // e si torna al tema scuro predefinito
        Assert.True(new FileInfo(Path.Combine(shots, "powers-scuro.png")).Length > 5000 && new FileInfo(Path.Combine(shots, "powers-chiaro.png")).Length > 5000);
    }

    static string Sha(string f) => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(f)));

    [Theory]
    [InlineData("fade", null)] [InlineData("push", "l")] [InlineData("wipe", "r")] [InlineData("cover", "u")] [InlineData("pull", "l")]
    public void TransitionPlayback_ShowsAnIntermediateFrame(string kind, string? dir)
    {
        var f = Env.Copy("mixed-realistic-1.pptx", "-play-" + kind);
        using (var d = Core.PptxDocument.Open(f)) { d.SetTransition(1, kind, dir, "slow"); d.SaveAs(f); }
        var prefix = Path.Combine(Env.Work, "play-" + kind);
        var exe = Path.Combine(Env.Repo, "src", "PowerPresentations.App", "bin", "Debug", "net10.0-windows", "PowerPresentations.App.exe");
        using var p = Process.Start(new ProcessStartInfo(exe, $"--transition-frames \"{f}\" 1 \"{prefix}\"") { UseShellExecute = false })!;
        Assert.True(p.WaitForExit(60000), "app non terminata"); Assert.Equal(0, p.ExitCode);
        string a = Sha(prefix + "-0-prima.png"), m = Sha(prefix + "-1-meta.png"), z = Sha(prefix + "-2-dopo.png");
        Assert.NotEqual(a, z);   // la slide cambia
        Assert.NotEqual(m, a);   // a metà animazione il fotogramma non è più la slide di partenza
        Assert.NotEqual(m, z);   // e non è ancora quella di arrivo: la transizione viene davvero riprodotta
    }

    [Fact]
    public void RealInput_UndoRedo_Buttons_And_CtrlZY_FileTab_CloseAsksToSave()
    {
        var json = Run("ui-undo.ps1", "").Trim().Split((char)10).Last().Trim();
        using var d = JsonDocument.Parse(json); var r = d.RootElement;
        Assert.Equal(4, r.GetProperty("afterDuplicate").GetInt32());
        Assert.Equal(3, r.GetProperty("afterUndo").GetInt32());                 // pulsante Annulla
        Assert.Equal(4, r.GetProperty("afterRedo").GetInt32());                 // pulsante Ripristina
        Assert.True(r.GetProperty("backstage").GetBoolean());                   // scheda File a tutta pagina
        Assert.Equal(3, r.GetProperty("afterCtrlZ").GetInt32());                // Ctrl+Z
        Assert.Equal(4, r.GetProperty("afterCtrlY").GetInt32());                // Ctrl+Y dopo Ctrl+Z (il fuoco non va perso)
        Assert.True(r.GetProperty("dialogShown").GetBoolean());                 // documento modificato: chiedere se salvare
        Assert.True(r.GetProperty("closed").GetBoolean());                      // risposta No: l'app si chiude
    }

    [Fact]
    public void WindowResize_CompactsRibbon_AndMaximizeButtonFollowsState_UiaOnly()
    {
        var shots = Directory.CreateDirectory(Path.Combine(Env.Work, "ui-resize")).FullName;
        var json = Run("ui-resize.ps1", $"-Shots \"{shots}\"").Trim().Split((char)10).Last().Trim();
        using var d = JsonDocument.Parse(json); var r = d.RootElement;
        Assert.Equal("Ingrandisci", r.GetProperty("maxNameNormal").GetString());                 // finestra: si può ingrandire
        Assert.Contains("Ripristina gi", r.GetProperty("maxNameMaximized").GetString());          // ingrandita: si può ripristinare
        Assert.Equal("Ingrandisci", r.GetProperty("maxNameBack").GetString());
        foreach (var n in new[] { "massimizzata", "1100", "800", "640" }) Assert.True(new FileInfo(Path.Combine(shots, $"resize-{n}.png")).Length > 8000, n);
        Assert.True(r.GetProperty("overflowButton").GetBoolean());                               // a 800 px compare "Altri comandi"
        Assert.Equal("Seleziona,Trova e sostituisci", r.GetProperty("overflowItems").GetString());   // il gruppo che non ci sta finisce nel menu
        Assert.StartsWith("800x", r.GetProperty("w640").GetString());                            // larghezza minima raggiunta (640 DIP a scala 125%)
    }

    [Fact]
    public void PresentationAndPresenterWindows_Render()
    {
        var o = Path.Combine(Env.Work, "ui-present.png");
        Run("screenshot.ps1", $"-File \"{Env.CorpusPath("notes-speaker.pptx")}\" -Out \"{o}\" -Mode present");
        Assert.True(new FileInfo(Path.Combine(Env.Work, "ui-present-presentazione.png")).Length > 5000);
        Assert.True(new FileInfo(Path.Combine(Env.Work, "ui-present-relatore.png")).Length > 5000);
    }

    [Fact]
    public void MainWindow_RendersCorpusFile_AndSlideSorterView()
    {
        Run("screenshot.ps1", $"-File \"{Env.CorpusPath("layouts-all.pptx")}\" -Out \"{Path.Combine(Env.Work, "ui-sorter.png")}\" -Mode sorter");
        Assert.True(new FileInfo(Path.Combine(Env.Work, "ui-sorter.png")).Length > 10000);
    }

    [Fact]
    public void RealInput_Commands_Clipboard_FindReplace_Bullets_Link_Grid_ApplyAll()
    {
        var work = Directory.CreateDirectory(Path.Combine(Env.Work, "ui-cmd")).FullName;
        string find, repl = "Progetto"; int n;
        using (var d = Core.PptxDocument.Open(Env.CorpusPath("mixed-realistic-1.pptx")))
        {
            n = d.ShapeIds(0).Count; Assert.True(n >= 2);
            find = new Core.Rendering.SlideRenderer(d).Render(0).Items.OfType<Core.Rendering.TextItem>().First().Text.Split(' ')[0];
        }
        var json = Run("ui-commands.ps1", $"-Find \"{find}\" -Repl \"{repl}\" -Work \"{work}\"", 240000).Trim().Split((char)10).Last().Trim();
        using var j = JsonDocument.Parse(json); var r = j.RootElement;
        Assert.False(r.TryGetProperty("error", out var err), json + " :: " + err);
        Assert.Contains("forme selezionate", r.GetProperty("afterSelectAll").GetString());                       // Ctrl+A reale
        int Shapes(string f) { using var d = Core.PptxDocument.Open(Path.Combine(work, f)); return d.ShapeIds(0).Count; }
        Assert.Equal(2 * n, Shapes("pasted.pptx"));                                                              // Ctrl+C, Ctrl+V
        Assert.Equal(n, Shapes("undone.pptx"));                                                                  // Ctrl+Z: un solo annulla
        Assert.Equal(0, Shapes("cut.pptx")); Assert.Equal(n, Shapes("cutpasted.pptx"));                          // Ctrl+X, Ctrl+V
        using (var d = Core.PptxDocument.Open(Path.Combine(work, "replaced.pptx")))
        {
            var all = string.Join("|", Enumerable.Range(0, d.Slides.Count).SelectMany(i => d.ShapeIds(i).Select(id => d.GetShapeText(i, id))));
            Assert.DoesNotContain(find, all); Assert.Contains(repl, all);                                        // Trova e sostituisci
        }
        using (var d = Core.PptxDocument.Open(Path.Combine(work, "shape.pptx")))
        {
            int first = d.ShapeIds(0)[0];
            Assert.Equal("bullet", d.GetBullet(0, first)); Assert.Equal("https://example.org/ui", d.GetHyperlink(0, first));   // Seleziona, Elenco puntato, Collegamento
        }
        foreach (var k in new[] { "Griglia", "Righello" }) { Assert.True(r.GetProperty(k + "-enabled").GetBoolean()); Assert.True(r.GetProperty(k + "-on").GetBoolean()); }
        using (var d = Core.PptxDocument.Open(Path.Combine(work, "trans.pptx"))) for (int i = 0; i < d.Slides.Count; i++) Assert.Equal("fade", d.GetTransition(i).Kind);   // Applica a tutte
        Assert.True(r.GetProperty("videoEnabled").GetBoolean());
    }

    [Fact]
    public void RealInput_CommentsPane_ListEditDelete_OneCtrlZ_SlideSwitch()
    {
        var work = Directory.CreateDirectory(Path.Combine(Env.Work, "ui-comments")).FullName; var shots = Directory.CreateDirectory(Path.Combine(work, "shots")).FullName;
        var src = Path.Combine(work, "in.pptx");
        using (var d = Core.PptxDocument.Open(Env.CorpusPath("mixed-realistic-1.pptx"))) { d.AddComment(0, "Primo"); d.AddComment(0, "Secondo"); CommentEditTests.AddReply(d, 0, "0", "Risposta"); d.SaveAs(src); }
        var orig = PptxComments(src);
        var json = Run("ui-comments.ps1", $"-Src \"{src}\" -Work \"{work}\" -Shots \"{shots}\"", 240000).Trim().Split((char)10).Last().Trim();
        using var j = JsonDocument.Parse(json); var r = j.RootElement;
        Assert.False(r.TryGetProperty("error", out var err), json + " :: " + err);
        Assert.False(r.GetProperty("paneBefore").GetBoolean());                                   // riquadro chiuso all'avvio
        Assert.Equal(3, r.GetProperty("listedButtons").GetInt32());                               // Commento apre il riquadro: commento, risposta, commento
        foreach (var t in new[] { "Primo", "Risposta", "Secondo", "Autore" }) Assert.Contains(t, r.GetProperty("listedTexts").GetString());
        var ed = PptxComments(Path.Combine(work, "edited.pptx"));                                 // Modifica con clic e tastiera reali
        Assert.Equal(new[] { "Primo", "Risposta", "Secondo modificato" }, ed.Select(c => c.Text));
        Assert.Equal(orig.Select(c => (c.Author, c.Created, c.Id)), ed.Select(c => (c.Author, c.Created, c.Id)));   // autore e data originali
        Assert.True(r.GetProperty("confirmShown").GetBoolean());                                  // conferma perché ha una risposta
        Assert.Equal(1, r.GetProperty("afterDeleteButtons").GetInt32());
        Assert.Equal(new[] { "Secondo modificato" }, PptxComments(Path.Combine(work, "deleted.pptx")).Select(c => c.Text));   // commento e risposta spariti
        Assert.Equal(3, r.GetProperty("afterUndoButtons").GetInt32());                            // un solo Ctrl+Z
        Assert.Equal(new[] { "Primo", "Risposta", "Secondo modificato" }, PptxComments(Path.Combine(work, "undone.pptx")).Select(c => c.Text));
        Assert.Equal(0, r.GetProperty("otherSlideButtons").GetInt32());                           // altra slide: il riquadro si aggiorna
        Assert.Contains("Nessun commento", r.GetProperty("otherSlideTexts").GetString());
        Assert.Equal(3, r.GetProperty("backButtons").GetInt32());
        foreach (var n in new[] { "commenti-riquadro", "commenti-dopo-elimina", "commenti-altra-slide" }) Assert.True(new FileInfo(Path.Combine(shots, n + ".png")).Length > 8000, n);
    }

    static IReadOnlyList<Core.CommentInfo> PptxComments(string f) { using var d = Core.PptxDocument.Open(f); return d.GetComments(0); }
}
