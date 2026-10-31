using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Tests;

public class RendererTests
{
    static DisplayList Render(string file, int slide = 0, RenderOptions? o = null)
    {
        using var d = PptxDocument.Open(Env.CorpusPath(file));
        return new SlideRenderer(d).Render(slide, o);
    }

    static IEnumerable<TextItem> Texts(DisplayList dl) => dl.Items.OfType<TextItem>();

    [Theory, MemberData(nameof(Env.CorpusFiles), MemberType = typeof(Env))]
    public void EverySlideOfCorpus_Renders_WithFiniteCoordinates(string name)
    {
        using var d = PptxDocument.Open(Env.CorpusPath(name));
        var r = new SlideRenderer(d);
        for (int i = 0; i < d.Slides.Count; i++)
        {
            var dl = r.Render(i);
            Assert.True(dl.Width > 0 && dl.Height > 0);
            foreach (var it in dl.Items)
                switch (it)
                {
                    case TextItem t: Assert.True(double.IsFinite(t.X) && double.IsFinite(t.Baseline) && t.Size > 0, $"{name} slide {i + 1}: testo con coordinate non valide"); break;
                    case PathItem p: Assert.All(p.Segs, s => Assert.True(double.IsFinite(s.X1) && double.IsFinite(s.Y1) && double.IsFinite(s.X3))); break;
                    case ImageItem im: Assert.True(im.Bounds.W >= 0 && im.Bounds.H >= 0); break;
                }
        }
    }

    [Fact]
    public void TitleSlide_Title_InheritsSizeCenteringAndFontFromMasterAndLayout()
    {
        var dl = Render("title-slide.pptx");
        var t = Texts(dl).Single(x => x.Text == "Titolo della presentazione");
        Assert.Equal(60, t.Size);                      // ctrTitle del layout "Diapositiva titolo": 6000
        var w = FontFace.Get(t.Font, t.Bold, t.Italic).Measure(t.Text, t.Size);
        Assert.InRange(t.X + w / 2, dl.Width / 2 - 2, dl.Width / 2 + 2); // centrato (algn=ctr ereditato dal layout)
        Assert.Contains("Aptos", t.Font);               // +mj-lt → font maggiore del tema
        var sub = Texts(dl).Single(x => x.Text == "Sottotitolo di prova");
        Assert.True(sub.Baseline > t.Baseline && sub.Size < t.Size);
    }

    [Fact]
    public void ThemeColors_AreReadFromFile()
    {
        var dl = Render("theme-custom-colors.pptx");
        var fills = dl.Items.OfType<PathItem>().Select(p => (p.Fill as SolidFill)?.Color).Where(c => c != null).ToList();
        Assert.Contains(new Rgba(200, 30, 30), fills); Assert.Contains(new Rgba(30, 160, 60), fills); Assert.Contains(new Rgba(30, 60, 200), fills);
    }

    [Fact]
    public void ThemeFonts_AreReadFromFile()
    {
        var dl = Render("theme-custom-fonts.pptx");
        Assert.Equal("Georgia", Texts(dl).Single(t => t.Text == "Font del tema").Font);
        Assert.Equal("Verdana", Texts(dl).First(t => t.Text.StartsWith("Testo in font minore")).Font);
    }

    [Fact]
    public void MasterBackgroundAndSlideBackgroundOverride()
    {
        Assert.Equal(new SolidFill(new Rgba(250, 240, 230)), Render("master-custom.pptx").Background);
        Assert.Equal(new SolidFill(new Rgba(30, 40, 80)), Render("slide-background.pptx").Background);
        Assert.IsType<GradientFill>(Render("slide-background.pptx", 1).Background);
    }

    [Fact]
    public void MasterShapes_AreDrawnUnderSlideContent()
    {
        var dl = Render("master-custom.pptx");
        var bar = dl.Items.OfType<PathItem>().First(p => p.Fill is SolidFill s && s.Color == new Rgba(160, 30, 30));
        Assert.True(bar.Segs.Max(s => s.Y1) > 500);
        Assert.True(dl.Hits.All(h => h.Id != bar.ShapeId)); // non selezionabile: appartiene al master
    }

    [Fact]
    public void LayoutPlaceholders_AreNotDrawnOnSlide_AndLayoutNamesRead()
    {
        using var d = PptxDocument.Open(Env.CorpusPath("layouts-all.pptx"));
        Assert.True(d.Layouts.Count >= 11);
        var dl = new SlideRenderer(d).Render(1);
        Assert.Contains(Texts(dl), t => t.Text == "Layout 2");
        Assert.DoesNotContain(Texts(dl), t => t.Text.Contains("Fare clic"));
    }

    [Fact]
    public void TextBox_WrapsInsideItsWidth()
    {
        var dl = Render("textbox-basic.pptx");
        var lines = Texts(dl).Where(t => t.Baseline > 250).ToList();
        Assert.True(lines.Count >= 3, "il testo lungo deve andare a capo");
        Assert.True(lines.Select(l => l.Baseline).Distinct().Count() == lines.Count);
        foreach (var l in lines) Assert.True(l.X + FontFace.Get(l.Font, l.Bold, l.Italic).Measure(l.Text, l.Size) <= 100 + 300 + 1);
    }

    [Fact]
    public void Bullets_AndNumbers_AreGenerated()
    {
        var dl = Render("bullets-numbered.pptx");
        Assert.Contains(Texts(dl), t => t.Text == "•"); Assert.Contains(Texts(dl), t => t.Text == "1."); Assert.Contains(Texts(dl), t => t.Text == "3.");
    }

    [Fact]
    public void Table_MergedCells_DrawnOnceWithAllText()
    {
        var dl = Render("table-merged-cells.pptx");
        Assert.Contains(Texts(dl), t => t.Text == "1,2"); Assert.Contains(Texts(dl), t => t.Text == "4,4"); Assert.Contains(Texts(dl), t => t.Text == "4,3");
        Assert.DoesNotContain(Texts(dl), t => t.Text == "1,4 x");
        Assert.Contains(dl.Items.OfType<PathItem>(), p => p.Fill is SolidFill s && s.Color == new Rgba(255, 230, 100)); // cella evidenziata
        // intestazione: testo bianco su accent1 (stile tabella)
        Assert.Equal(new Rgba(255, 255, 255), Texts(dl).First(t => t.Text == "1,4").Color);
    }

    [Fact]
    public void Picture_IsLoaded_WithCrop()
    {
        var dl = Render("picture-crop-rotate.pptx");
        var imgs = dl.Items.OfType<ImageItem>().ToList();
        Assert.Equal(2, imgs.Count);
        Assert.All(imgs, i => { Assert.NotNull(i.Data); Assert.Equal("image/png", i.ContentType); });
        Assert.True(imgs[0].Crop.X > 0 && imgs[0].Crop.Y > 0);
        Assert.Equal(15, imgs[1].Rot);
    }

    [Fact]
    public void Group_ChildrenMappedToSlideSpace()
    {
        var dl = Render("group-shapes.pptx");
        var paths = dl.Items.OfType<PathItem>().ToList();
        Assert.True(paths.Count >= 2);
        Assert.All(paths, p => Assert.All(p.Segs.Where(s => s.Kind != SegKind.Close), s => Assert.InRange(s.X1, 380, 800)));
    }

    [Fact]
    public void Rotation_AndFlip_ChangeGeometry()
    {
        var dl = Render("shape-rotation-flip.pptx");
        var rot = dl.Items.OfType<PathItem>().First();
        Assert.Contains(rot.Segs, s => s.Kind == SegKind.Line && Math.Abs(s.Y1 - rot.Segs[0].Y1) > 20); // arrow ruotata di 30°: non più orizzontale
    }

    [Fact]
    public void Autofit_StoredScaleIsUsed_AndRecomputeFindsFit()
    {
        var plain = Render("autofit-shrink.pptx");
        var sz = Texts(plain).Where(t => t.Text.StartsWith("Punto")).Max(t => t.Size);
        var fit = Render("autofit-shrink.pptx", 0, new RenderOptions { RecomputeAutofit = true });
        Assert.True(fit.Fits.Values.Min(f => f.Scale) < 1);
        Assert.True(Texts(fit).Where(t => t.Text.StartsWith("Punto")).Max(t => t.Size) < sz);
        var h = Texts(fit).Where(t => t.Text.StartsWith("Punto")).Max(t => t.Baseline);
        Assert.True(h <= plain.Height, "con il ricalcolo il testo sta nella slide");
    }

    [Fact]
    public void Transitions_AreRead()
    {
        using var d = PptxDocument.Open(Env.CorpusPath("transitions.pptx"));
        Assert.Equal("fade", d.GetTransition(0).Kind);
        Assert.All(Enumerable.Range(0, d.Slides.Count), i => Assert.NotEqual("none", d.GetTransition(i).Kind));
    }

    [Fact]
    public void ColorTransforms_TintIsLinearLikeOffice()
    {
        var cc = new ColorCtx(new ThemeInfo(), ColorCtx.DefaultMap);
        var c = cc.Resolve(System.Xml.Linq.XElement.Parse("<a:srgbClr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" val=\"156082\"><a:tint val=\"40000\"/></a:srgbClr>"))!.Value;
        Assert.InRange(c.R, 203, 205); Assert.InRange(c.G, 209, 211); Assert.InRange(c.B, 215, 217); // pixel misurati nel PDF esportato da PowerPoint: (204,210,216)
    }
}
