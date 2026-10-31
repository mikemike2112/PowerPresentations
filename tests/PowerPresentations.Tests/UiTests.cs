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
