using System.Diagnostics;
using System.Text.Json;

namespace PowerPresentations.Tests;

/// <summary>Prove sulla UI reale (WPF) avviata come processo: UI Automation, input reale del mouse/tastiera, screenshot. Solo processi PowerPresentations.App avviati dal test.</summary>
[Collection("Ui"), Trait("Category", "Ui")]
public class UiTests
{
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
}

[CollectionDefinition("Ui", DisableParallelization = true)]
public class UiCollection { }
