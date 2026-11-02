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

    // ---- copia / taglia / incolla ----
    [Fact]
    public void CopyPaste_PictureAndText_BetweenSlides_ValidAndImageBytesKept()
    {
        using var d = PptxDocument.New();
        int pic = d.AddPicture(0, ComEditFixture.Png, 100, 100, 50, 50), tb = d.AddTextBox(0, 10, 10, 200, 30, "Ciao");
        var clip = ShapeClip.FromBytes(d.CopyShapes(0, new[] { pic, tb }).ToBytes())!;   // passa dal formato della clipboard
        Assert.Equal("Ciao", clip.Text);
        var s2 = d.AddSlide(d.Layouts[1], 0);
        var ids = d.PasteShapes(s2, clip);
        Assert.Equal(2, ids.Count); Assert.Equal(2, ids.Distinct().Count());
        d.PasteShapes(s2, clip);   // secondo incolla nello stesso punto: spostato, id distinti
        using var r = Reload(d, "paste");
        Assert.Equal(2, r.Slides[s2].Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Count());
        var imgs = r.Slides[s2].ImageParts.ToList(); Assert.NotEmpty(imgs);
        foreach (var ip in imgs) { using var st = ip.GetStream(); var ms = new MemoryStream(); st.CopyTo(ms); Assert.Equal(ComEditFixture.Png, ms.ToArray()); }
        var offs = r.Slides[s2].Slide!.Descendants<DocumentFormat.OpenXml.Drawing.Offset>().Select(o => (o.X!.Value, o.Y!.Value)).ToList();
        Assert.Equal(offs.Count, offs.Distinct().Count());
        var all = r.Slides[s2].Slide!.Descendants<DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties>().Select(n => n.Id!.Value).ToList();
        Assert.Equal(all.Count, all.Distinct().Count());
    }

    [Fact]
    public void CopyPaste_Placeholder_BecomesFreeShape_WithExplicitBox()
    {
        using var d = PptxDocument.New();
        int title = FirstText(d, 0); d.SetShapeText(0, title, "Titolo");
        var clip = d.CopyShapes(0, new[] { title });
        var s2 = d.AddSlide(d.Layouts[1], 0); d.PasteShapes(s2, clip);
        using var r = Reload(d, "paste-ph");
        var pasted = r.Slides[s2].Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Last();
        Assert.Null(pasted.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.GetFirstChild<DocumentFormat.OpenXml.Presentation.PlaceholderShape>());
        Assert.NotNull(pasted.ShapeProperties!.Transform2D);
    }

    [Fact]
    public void CopyPaste_Chart_AcrossDocuments_KeepsChartPartAndWorkbook()
    {
        using var src = PptxDocument.Open(Env.CorpusPath("chart-office.pptx"));
        int gf = src.ShapeIds(0).First(i => src.ShapeLabel(0, i).Contains("Grafico"));
        var clip = src.CopyShapes(0, new[] { gf });
        using var dst = PptxDocument.New(); dst.PasteShapes(0, clip);
        var f = Out("paste-chart"); dst.SaveAs(f); Assert.Empty(RoundTripTests.NewErrors(Env.CorpusPath("chart-office.pptx"), f));
        using var z = System.IO.Compression.ZipFile.OpenRead(f);
        Assert.Contains(z.Entries, e => e.FullName.Contains("/chart") && e.FullName.EndsWith(".xml") && !e.FullName.Contains("_rels"));
        Assert.Contains(z.Entries, e => e.FullName.EndsWith(".xlsx"));
    }

    [Fact]
    public void Paste_RejectsGarbage_AndUnsafeLinkSchemes()
    {
        Assert.Null(ShapeClip.FromBytes(new byte[] { 1, 2, 3 })); Assert.Null(ShapeClip.FromBytes(null));
        using var d = PptxDocument.New();
        var bad = new ShapeClip { Shapes = { "<p:script xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>" } };
        Assert.Throws<PptxFormatException>(() => d.PasteShapes(0, bad));
        var xml = "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:nvSpPr><p:cNvPr id=\"2\" name=\"x\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id=\"rId9\"/></a:rPr><a:t>l</a:t></a:r></a:p></p:txBody></p:sp>";
        var evil = new ShapeClip { Shapes = { xml }, Parts = { ["rId9"] = new ClipPart { Kind = "hyper", RelType = "hyperlink", Uri = "file:///c:/windows/system32/calc.exe", External = true } } };
        d.PasteShapes(0, evil);
        using var r = Reload(d, "paste-evil");
        Assert.DoesNotContain("calc.exe", string.Join("", r.Slides[0].HyperlinkRelationships.Select(h => h.Uri.ToString())));
        Assert.Empty(r.Slides[0].Slide!.Descendants<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>());
    }

    // ---- elenchi ----
    [Fact]
    public void Bullets_Set_Toggle_AndKeepPprChildOrder_Valid()
    {
        var src = Env.CorpusPath("paragraph-align-spacing.pptx");
        using var d = PptxDocument.Open(src);
        int id = FirstText(d, 0);
        d.SetBullet(0, id, "bullet"); Assert.Equal("bullet", d.GetBullet(0, id));
        d.SetBullet(0, id, "number"); Assert.Equal("number", d.GetBullet(0, id));
        d.SetBullet(0, id, "none"); Assert.Equal("none", d.GetBullet(0, id));
        d.SetBullet(0, id, "number");
        var f = Out("bullets"); d.SaveAs(f); Assert.Empty(RoundTripTests.NewErrors(src, f));
        using var r = PptxDocument.Open(f); Assert.Equal("number", r.GetBullet(0, id));
        Assert.Throws<ArgumentException>(() => r.SetBullet(0, id, "x"));
    }

    [Fact]
    public void Bullets_OnTable_AndPicture()
    {
        using var d = PptxDocument.New();
        int t = d.AddTable(0, 2, 2, 50, 50, 300), pic = d.AddPicture(0, ComEditFixture.Png, 10, 10, 20, 20);
        d.SetBullet(0, t, "bullet"); d.SetBullet(0, pic, "bullet"); Assert.Equal("bullet", d.GetBullet(0, t)); Assert.Equal("none", d.GetBullet(0, pic));
        using var r = Reload(d, "bullets-table");
    }

    // ---- applica a tutte ----
    [Fact]
    public void ApplyTransitionToAll_CopiesAlternateContentAndPlainTransitions()
    {
        var src = Env.CorpusPath("transitions.pptx");
        using var d = PptxDocument.Open(src);
        d.SetTransition(0, "push", "l", "fast");
        d.ApplyTransitionToAll(0);
        for (int i = 0; i < d.Slides.Count; i++) { Assert.Equal("push", d.GetTransition(i).Kind); Assert.Equal("fast", d.GetTransition(i).Speed); }
        var f = Out("trans-all"); d.SaveAs(f); Assert.Empty(RoundTripTests.NewErrors(src, f));
        using var r = PptxDocument.Open(f);
        for (int i = 0; i < r.Slides.Count; i++) Assert.Equal("push", r.GetTransition(i).Kind);
        r.SetTransition(0, "none"); r.ApplyTransitionToAll(0);
        for (int i = 0; i < r.Slides.Count; i++) Assert.Equal("none", r.GetTransition(i).Kind);
    }
}
