using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Tests;

public class ReviewCommandTests
{
    static string Out(string n) => Path.Combine(Env.Work, "rev-" + n + ".pptx");

    [Fact]
    public void Comment_Modern_AddedToNewAndExistingFile_ReadBack_NoPersonalName()
    {
        using var d = PptxDocument.New();
        d.AddComment(0, "Primo commento"); d.AddComment(0, "Secondo\nsu due righe");
        var c = d.GetComments(0); Assert.Equal(2, c.Count); Assert.Equal("Autore", c[0].Author); Assert.Equal("Primo commento", c[0].Text);
        Assert.Throws<ArgumentException>(() => d.AddComment(0, "  "));
        using var r = CommandTests.Reload(d, "comment");
        Assert.Equal(2, r.GetComments(0).Count);
        var f = Path.Combine(Env.Work, "cmd-comment.pptx"); using var z = System.IO.Compression.ZipFile.OpenRead(f);
        foreach (var e in z.Entries.Where(e => e.FullName.Contains("author") || e.FullName.Contains("comment"))) Assert.DoesNotContain(Environment.UserName, new StreamReader(e.Open()).ReadToEnd());

        var src = Env.CorpusPath("comments.pptx");   // file con un commento moderno già presente
        using var e2 = PptxDocument.Open(src); int before = e2.GetComments(0).Count;
        e2.AddComment(0, "Aggiunto"); Assert.Equal(before + 1, e2.GetComments(0).Count);
        var f2 = Out("comment-existing"); e2.SaveAs(f2); Assert.Empty(RoundTripTests.NewErrors(src, f2));
        using var r2 = PptxDocument.Open(f2); Assert.Equal(before + 1, r2.GetComments(0).Count);
    }

    [Fact]
    public void HeaderFooter_AddUpdateRemove_AllSlides_Valid()
    {
        using var d = PptxDocument.New();
        d.AddSlide(d.Layouts[1], 0);
        d.ApplyHeaderFooter(new HeaderFooter(true, true, false, "02/11/2026", true, "Piè di prova & co"));
        for (int i = 0; i < 2; i++) { var h = d.GetHeaderFooter(i); Assert.True(h.SlideNumber && h.Date && h.Footer); Assert.Equal("Piè di prova & co", h.FooterText); Assert.Equal("02/11/2026", h.DateText); Assert.False(h.DateAuto); }
        var texts = new SlideRenderer(d).Render(1).Items.OfType<TextItem>().Select(t => t.Text).ToList();
        Assert.Contains("2", texts); Assert.Contains("Piè di prova & co", texts);
        d.ApplyHeaderFooter(new HeaderFooter(false, true, true, "", false, ""), 1);   // solo slide 2: data automatica, niente numero né piè
        Assert.True(d.GetHeaderFooter(1).DateAuto); Assert.False(d.GetHeaderFooter(1).SlideNumber); Assert.True(d.GetHeaderFooter(0).SlideNumber);
        d.ApplyHeaderFooter(new HeaderFooter(true, false, false, "", true, "Nuovo"), 0);   // aggiorna senza duplicare
        Assert.Equal(1, d.Slides[0].Slide!.Descendants<DocumentFormat.OpenXml.Presentation.PlaceholderShape>().Count(p => p.Type?.Value == DocumentFormat.OpenXml.Presentation.PlaceholderValues.Footer));
        using var r = CommandTests.Reload(d, "hf");
        Assert.Equal("Nuovo", r.GetHeaderFooter(0).FooterText);
    }

    [Fact]
    public void HeaderFooter_SkipTitleLayout_AndLayoutWithoutPlaceholders()
    {
        using var d = PptxDocument.New();
        d.ApplyHeaderFooter(new HeaderFooter(true, false, false, "", false, "", true));   // slide 1 usa il layout titolo
        Assert.False(d.GetHeaderFooter(0).SlideNumber);
        var src = Env.CorpusPath("footer-slidenumber.pptx"); using var f = PptxDocument.Open(src);
        Assert.True(f.GetHeaderFooter(0).SlideNumber); Assert.Equal("Piè di pagina di prova", f.GetHeaderFooter(0).FooterText);
        f.ApplyHeaderFooter(new HeaderFooter(false, false, false, "", false, ""));
        var o = Out("hf-corpus"); f.SaveAs(o); Assert.Empty(RoundTripTests.NewErrors(src, o));
        Assert.Throws<ArgumentException>(() => f.ApplyHeaderFooter(new HeaderFooter(false, false, false, "", true, new string('x', 501))));
    }

    [Fact]
    public void ShowSettings_RoundTrip_AndValidation()
    {
        using var d = PptxDocument.New(); d.AddSlide(d.Layouts[1], 0); d.AddSlide(d.Layouts[1], 1);
        Assert.Equal(new ShowSettings("present", false, true, true, true, null, null), d.GetShowSettings());
        var s = new ShowSettings("kiosk", true, false, false, false, 2, 3); d.SetShowSettings(s);
        Assert.Equal(s, d.GetShowSettings());
        Assert.Throws<ArgumentException>(() => d.SetShowSettings(s with { From = 3, To = 2 })); Assert.Throws<ArgumentException>(() => d.SetShowSettings(s with { To = 9 })); Assert.Throws<ArgumentException>(() => d.SetShowSettings(s with { Mode = "x" }));
        Assert.Equal(s, d.GetShowSettings());   // il rifiuto non cambia nulla
        using var r = CommandTests.Reload(d, "show"); Assert.Equal(s, r.GetShowSettings());
        r.SetShowSettings(new ShowSettings("browse", false, true, true, true, null, null)); Assert.Equal("browse", r.GetShowSettings().Mode);
        var src = Env.CorpusPath("mixed-realistic-1.pptx"); using var m = PptxDocument.Open(src); m.SetShowSettings(s with { To = 3 });
        var o = Out("show-corpus"); m.SaveAs(o); Assert.Empty(RoundTripTests.NewErrors(src, o));
    }
}
