using PowerPresentations.Core;

namespace PowerPresentations.Tests;

public class TimingTests
{
    static string Out(string n) => Path.Combine(Env.Work, "tim-" + n + ".pptx");

    [Fact]
    public void AdvanceTime_SetOnPlainAndAlternateContentTransitions_KeptWhenEffectChanges_Valid()
    {
        var src = Env.CorpusPath("transitions.pptx");
        using var d = PptxDocument.Open(src);
        Assert.Null(d.GetAdvanceTime(0));
        for (int i = 0; i < d.Slides.Count; i++) d.SetAdvanceTime(i, 1500 + i);
        for (int i = 0; i < d.Slides.Count; i++) Assert.Equal(1500 + i, d.GetAdvanceTime(i));
        var kinds = Enumerable.Range(0, d.Slides.Count).Select(i => d.GetTransition(i).Kind).ToList();
        d.SetTransition(1, "push", "l", "fast"); Assert.Equal(1501, d.GetAdvanceTime(1)); Assert.Equal("push", d.GetTransition(1).Kind);
        d.SetTransition(1, "none"); Assert.Equal(1501, d.GetAdvanceTime(1));   // senza effetto resta l'intervallo
        d.SetAdvanceTime(1, null); Assert.Null(d.GetAdvanceTime(1)); Assert.Equal("none", d.GetTransition(1).Kind);
        var f = Out("adv"); d.SaveAs(f); Assert.Empty(RoundTripTests.NewErrors(src, f));
        using var r = PptxDocument.Open(f); Assert.Equal(1500, r.GetAdvanceTime(0)); Assert.Equal(1502, r.GetAdvanceTime(2)); Assert.Null(r.GetAdvanceTime(1));
        Assert.Throws<ArgumentOutOfRangeException>(() => r.SetAdvanceTime(0, -5));
    }

    [Fact]
    public void AdvanceTime_OnSlideWithoutTransition_AndApplyAll()
    {
        using var d = PptxDocument.New(); d.AddSlide(d.Layouts[1], 0);
        d.SetAdvanceTime(0, 4000); Assert.Equal(4000, d.GetAdvanceTime(0)); Assert.Null(d.GetAdvanceTime(1));
        d.ApplyTransitionToAll(0); Assert.Equal(4000, d.GetAdvanceTime(1));
        using var r = CommandTests.Reload(d, "adv-new"); Assert.Equal(4000, r.GetAdvanceTime(1));
        r.SetAdvanceTime(0, null); r.SetAdvanceTime(1, null);
        Assert.True(!r.Slides[0].Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Transition>().Any(), string.Join("|", r.Slides[0].Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Transition>().Select(t => t.OuterXml)));   // nessun elemento vuoto lasciato
    }
}
