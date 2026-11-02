using PowerPresentations.Core;

namespace PowerPresentations.Tests;

public class CustomShowTests
{
    [Fact]
    public void CustomShow_CreateReplaceDelete_FollowsSlideDeletion_Valid()
    {
        var src = Env.CorpusPath("mixed-realistic-1.pptx");
        using var d = PptxDocument.Open(src);
        d.SetCustomShow("Breve", new[] { 2, 0 }); d.SetCustomShow("Tutto", new[] { 0, 1, 2 });
        Assert.Equal(new[] { 2, 0 }, d.CustomShows.First(c => c.Name == "Breve").Slides);
        d.SetCustomShow("breve", new[] { 1 });   // stesso nome (senza maiuscole): sostituisce
        Assert.Equal(2, d.CustomShows.Count); Assert.Equal(new[] { 1 }, d.CustomShows.First(c => c.Name == "breve").Slides);
        Assert.Throws<ArgumentException>(() => d.SetCustomShow(" ", new[] { 0 })); Assert.Throws<ArgumentException>(() => d.SetCustomShow("x", Array.Empty<int>())); Assert.Throws<ArgumentOutOfRangeException>(() => d.SetCustomShow("x", new[] { 9 }));
        var f = Path.Combine(Env.Work, "cs.pptx"); d.SaveAs(f); Assert.Empty(RoundTripTests.NewErrors(src, f));
        using var r = PptxDocument.Open(f);
        r.DeleteSlide(0);   // le slide eliminate escono da tutte le presentazioni personalizzate
        Assert.Equal(new[] { 0, 1 }, r.CustomShows.First(c => c.Name == "Tutto").Slides);
        r.DeleteCustomShow("Tutto"); r.DeleteCustomShow("breve"); Assert.Empty(r.CustomShows); Assert.Throws<KeyNotFoundException>(() => r.DeleteCustomShow("Tutto"));
        var f2 = Path.Combine(Env.Work, "cs2.pptx"); r.SaveAs(f2); Assert.Empty(RoundTripTests.NewErrors(src, f2));
    }
}
