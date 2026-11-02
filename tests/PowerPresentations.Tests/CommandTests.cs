using DocumentFormat.OpenXml.Packaging;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Tests;

/// <summary>Comandi del ribbon: ogni test applica il comando, salva, valida con l'Open XML SDK e rilegge.</summary>
public class CommandTests
{
    static string Out(string n) => Path.Combine(Env.Work, "cmd-" + n + ".pptx");
    static int FirstText(PptxDocument d, int s) => new SlideRenderer(d).Render(s).Hits.First(h => h.HasText).Id;

    /// <summary>Salva, controlla che il validatore non trovi errori e riapre il file.</summary>
    internal static PptxDocument Reload(PptxDocument d, string name)
    {
        var f = Out(name); d.SaveAs(f);
        Assert.Empty(RoundTripTests.Errors(f));
        return PptxDocument.Open(f);
    }

    [Fact]
    public void Format_OnTable_AppliesToAllCells_AndOnPicture_DoesNothing()
    {
        using var d = PptxDocument.New();
        int t = d.AddTable(0, 2, 2, 50, 50, 300), pic = d.AddPicture(0, ComEditFixture.Png, 10, 10, 20, 20);
        d.SetCellText(0, t, 0, 0, "a");
        Assert.Equal(4, d.TextBodies(0, t).Count); Assert.Empty(d.TextBodies(0, pic));
        d.SetFormat(0, t, bold: true, sizePt: 20); d.SetAlignment(0, t, "ctr");
        d.SetFormat(0, pic, bold: true);   // nessuna eccezione
        Assert.Throws<KeyNotFoundException>(() => d.TextBodies(0, 9999));
        using var r = Reload(d, "tablefmt");
        var x = string.Concat(r.Slides[0].Slide!.Descendants<DocumentFormat.OpenXml.Drawing.Run>().Select(e => e.OuterXml));
        Assert.Contains("sz=\"2000\"", x); Assert.Contains("b=\"1\"", x);
        Assert.Equal((false, 20.0), r.GetRunState(0, t));
    }

    [Fact]
    public void RunState_ReadsUnderlineFromFile()
    {
        using var d = PptxDocument.New();
        int id = d.AddTextBox(0, 10, 10, 200, 30, "x");
        Assert.False(d.GetRunState(0, id).Underline);
        d.SetFormat(0, id, underline: true); Assert.True(d.GetRunState(0, id).Underline);
        d.SetFormat(0, id, underline: false); Assert.False(d.GetRunState(0, id).Underline);
    }
}
