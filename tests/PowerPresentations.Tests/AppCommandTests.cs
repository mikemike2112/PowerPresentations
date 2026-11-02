using System.Diagnostics;
using System.IO.Compression;

namespace PowerPresentations.Tests;

/// <summary>Comandi dell'app provati dalla riga di comando, senza mouse né tastiera reali (quindi senza blocco UI).</summary>
[Collection("Desktop"), Trait("Category", "Ui")]
public class AppCommandTests
{
    static string Exe => Path.Combine(Env.Repo, "src", "PowerPresentations.App", "bin", "Debug", "net10.0-windows", "PowerPresentations.App.exe");

    [Fact]
    public void Print_WritesOnePagePerVisibleSlide_ToXps()
    {
        Assert.True(File.Exists(Exe), "Compilare prima l'app: dotnet build");
        var src = Env.CorpusPath("hidden-slide.pptx"); var xps = Path.Combine(Env.Work, "print.xps");
        using var p = Process.Start(new ProcessStartInfo(Exe, $"--print-xps \"{src}\" \"{xps}\"") { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true })!;
        var so = p.StandardOutput.ReadToEndAsync(); var se = p.StandardError.ReadToEndAsync();
        Assert.True(p.WaitForExit(60000), "app non terminata"); Assert.True(p.ExitCode == 0, se.Result);
        using var d = Core.PptxDocument.Open(src);
        int visible = Enumerable.Range(0, d.Slides.Count).Count(i => !d.IsHidden(i)); Assert.True(visible < d.Slides.Count);   // il file ha una slide nascosta
        using var z = ZipFile.OpenRead(xps);
        Assert.Equal(visible, z.Entries.Count(e => e.FullName.EndsWith(".fpage", StringComparison.OrdinalIgnoreCase)));
    }
}
