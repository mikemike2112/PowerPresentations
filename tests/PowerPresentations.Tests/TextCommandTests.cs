using DocumentFormat.OpenXml.Packaging;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerPresentations.Tests;

public class TextCommandTests
{
    static string Out(string n) => Path.Combine(Env.Work, "txt-" + n + ".pptx");
    static int FirstText(PptxDocument d, int s) => new SlideRenderer(d).Render(s).Hits.First(h => h.HasText).Id;

    [Fact]
    public void FindReplace_CaseOptions_KeepsRunFormat_AndTables_Valid()
    {
        var src = Env.CorpusPath("text-formatting.pptx");
        using var d = PptxDocument.Open(src);
        int id = FirstText(d, 0); d.SetShapeText(0, id, "Alfa beta ALFA"); d.SetFormat(0, id, bold: true);
        int t = d.AddTable(0, 1, 2, 20, 20, 200); d.SetCellText(0, t, 0, 1, "alfa in cella");
        Assert.Equal(2, d.FindAll("alfa").Count); Assert.Single(d.FindAll("alfa", true), h => h.ShapeId == t);
        Assert.Throws<ArgumentException>(() => d.FindAll(""));
        Assert.Equal(3, d.ReplaceAll("alfa", "Omega")); Assert.Empty(d.FindAll("alfa"));
        Assert.Equal("Omega beta Omega", d.GetShapeText(0, id)); Assert.Equal("Omega in cella", d.GetCellText(0, t, 0, 1));
        Assert.Equal(0, d.ReplaceAll("zzzz", "x"));
        var f = Out("replace"); d.SaveAs(f); Assert.Empty(RoundTripTests.NewErrors(src, f));
        using var r = PptxDocument.Open(f);
        Assert.Contains("b=\"1\"", r.Slides[0].Slide!.OuterXml);
    }

    [Fact]
    public void Hyperlink_SetReuseRemove_RejectsUnsafeSchemes()
    {
        using var d = PptxDocument.New();
        int a = d.AddTextBox(0, 10, 10, 200, 30, "uno"), b = d.AddTextBox(0, 10, 60, 200, 30, "due");
        d.SetHyperlink(0, a, "example.org/x"); d.SetHyperlink(0, b, "https://example.org/x");
        Assert.Equal("https://example.org/x", d.GetHyperlink(0, a)); Assert.Single(d.Slides[0].HyperlinkRelationships);   // una sola relazione condivisa
        foreach (var bad in new[] { "file:///c:/windows/notepad.exe", "javascript:alert(1)", "", "ftp://x.org"})
            if (bad.Length > 0) Assert.Throws<ArgumentException>(() => d.SetHyperlink(0, a, bad));
        Assert.Equal("https://example.org/x", d.GetHyperlink(0, a));   // il rifiuto non tocca nulla
        d.SetHyperlink(0, a, "mailto:x@example.org");
        using var r = CommandTests.Reload(d, "link");
        Assert.Equal("mailto:x@example.org", r.GetHyperlink(0, a));
        r.SetHyperlink(0, a, null); r.SetHyperlink(0, b, ""); Assert.Null(r.GetHyperlink(0, a)); Assert.Empty(r.Slides[0].HyperlinkRelationships);
    }

    [Fact]
    public void AppendText_AddsSymbolInLastRun_AndInEmptyBox()
    {
        using var d = PptxDocument.New();
        int a = d.AddTextBox(0, 10, 10, 200, 30, "Prezzo 5"), e = d.AddTextBox(0, 10, 60, 200, 30);
        d.AppendText(0, a, "€"); d.AppendText(0, e, "©");
        Assert.Equal("Prezzo 5€", d.GetShapeText(0, a)); Assert.Equal("©", d.GetShapeText(0, e));
        Assert.Throws<NotSupportedException>(() => d.AppendText(0, d.AddPicture(0, ComEditFixture.Png, 1, 1, 5, 5), "x"));
        using var r = CommandTests.Reload(d, "symbol");
    }

    [Fact]
    public void CopyFormat_CopiesRunParagraphFillAndLine_KeepsTextAndLinks()
    {
        var src = Env.CorpusPath("shape-fills-lines.pptx");
        using var d = PptxDocument.Open(src);
        int s1 = d.AddShape(0, "rect", 10, 10, 100, 50, "origine"), s2 = d.AddShape(0, "ellipse", 200, 10, 100, 50, "destinazione");
        d.SetFormat(0, s1, bold: true, sizePt: 28, colorHex: "FF0000"); d.SetAlignment(0, s1, "r"); d.SetBullet(0, s1, "bullet");
        d.SetHyperlink(0, s2, "https://example.org");
        var fc = d.CopyFormat(0, s1); d.ApplyFormat(0, s2, fc);
        Assert.Equal("destinazione", d.GetShapeText(0, s2)); Assert.Equal("https://example.org", d.GetHyperlink(0, s2));
        Assert.Equal((false, 28.0), d.GetRunState(0, s2)); Assert.Equal("bullet", d.GetBullet(0, s2));
        var f = Out("copyfmt"); d.SaveAs(f); Assert.Empty(RoundTripTests.NewErrors(src, f));
        using var r = PptxDocument.Open(f); Assert.Equal((false, 28.0), r.GetRunState(0, s2));
        // su una tabella e su un'immagine non fallisce
        int t = r.AddTable(0, 1, 1, 5, 5, 50), pic = r.AddPicture(0, ComEditFixture.Png, 1, 1, 5, 5);
        r.ApplyFormat(0, t, fc); r.ApplyFormat(0, pic, fc);
    }
}
