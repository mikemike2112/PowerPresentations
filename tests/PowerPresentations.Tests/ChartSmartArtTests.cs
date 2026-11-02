using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Tests;

/// <summary>Grafico (con workbook incorporato) e SmartArt creati con l'SDK: devono sopravvivere a modifica+salvataggio e disegnarsi come segnaposto.</summary>
public class ChartSmartArtTests
{
    const string C = "http://schemas.openxmlformats.org/drawingml/2006/chart", D = "http://schemas.openxmlformats.org/drawingml/2006/diagram", A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    static MemoryStream S(string x) => new(Encoding.UTF8.GetBytes(x));

    static string Frame(int id, string name, string uri, string inner) =>
        $"<p:graphicFrame xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"{A}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:nvGraphicFramePr><p:cNvPr id=\"{id}\" name=\"{name}\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"4572000\" cy=\"2743200\"/></p:xfrm><a:graphic><a:graphicData uri=\"{uri}\">{inner}</a:graphicData></a:graphic></p:graphicFrame>";

    public static string Build()
    {
        var path = Path.Combine(Env.Work, "chart-smartart.pptx");
        File.Copy(Env.CorpusPath("mixed-realistic-1.pptx"), path, true);
        using (var d = PresentationDocument.Open(path, true))
        {
            var slide = d.PresentationPart!.SlideParts.First();
            var chart = slide.AddNewPart<ChartPart>("rIdChart");
            chart.FeedData(S($"<c:chartSpace xmlns:c=\"{C}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><c:chart><c:plotArea><c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:val><c:numLit><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>3</c:v></c:pt><c:pt idx=\"1\"><c:v>5</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart></c:plotArea></c:chart><c:externalData r:id=\"rIdXlsx\"><c:autoUpdate val=\"0\"/></c:externalData></c:chartSpace>"));
            var emb = chart.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rIdXlsx");
            emb.FeedData(new MemoryStream(new byte[] { 0x50, 0x4B, 3, 4, 1, 2, 3, 4, 5 }));
            var dm = slide.AddNewPart<DiagramDataPart>("rIdDm"); dm.FeedData(S($"<dgm:dataModel xmlns:dgm=\"{D}\"><dgm:ptLst/></dgm:dataModel>"));
            var lo = slide.AddNewPart<DiagramLayoutDefinitionPart>("rIdLo"); lo.FeedData(S($"<dgm:layoutDef xmlns:dgm=\"{D}\" uniqueId=\"urn:test/layout\"/>"));
            var qs = slide.AddNewPart<DiagramStylePart>("rIdQs"); qs.FeedData(S($"<dgm:styleDef xmlns:dgm=\"{D}\" uniqueId=\"urn:test/style\"/>"));
            var cs = slide.AddNewPart<DiagramColorsPart>("rIdCs"); cs.FeedData(S($"<dgm:colorsDef xmlns:dgm=\"{D}\" uniqueId=\"urn:test/colors\"/>"));
            var sp = slide.Slide.CommonSlideData!.ShapeTree!;
            sp.AppendChild(new DocumentFormat.OpenXml.Presentation.GraphicFrame(Frame(901, "Grafico prova", C, $"<c:chart xmlns:c=\"{C}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdChart\"/>")));
            sp.AppendChild(new DocumentFormat.OpenXml.Presentation.GraphicFrame(Frame(902, "SmartArt prova", D, $"<dgm:relIds xmlns:dgm=\"{D}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:dm=\"rIdDm\" r:lo=\"rIdLo\" r:qs=\"rIdQs\" r:cs=\"rIdCs\"/>")));
        }
        return path;
    }

    [Fact]
    public void ChartAndSmartArt_SurviveEditAndSave_AndRenderAsPlaceholder()
    {
        var f = Build(); var outp = Path.Combine(Env.Work, "chart-smartart-out.pptx");
        using (var d = PptxDocument.Open(f))
        {
            var r = new SlideRenderer(d).Render(1);
            Assert.Contains(r.Hits, h => h.Id == 901); Assert.Contains(r.Hits, h => h.Id == 902);   // disegnati (segnaposto), non ignorati
            var id = r.Hits.First(h => h.HasText).Id;
            d.SetShapeText(1, id, "Modificato"); d.MoveSlide(1, 0); d.SetNotes(0, "nota");
            d.SaveAs(outp);
        }
        using var z = System.IO.Compression.ZipFile.OpenRead(outp);
        var names = z.Entries.Select(e => e.FullName).ToList();
        foreach (var p in new[] { "ppt/slides/charts/chart", "ppt/slides/charts/embeddings/", "ppt/graphics/data", "ppt/graphics/layout", "ppt/graphics/quickStyle", "ppt/graphics/colors" })
            Assert.True(names.Any(n => n.StartsWith(p)), p + " in " + string.Join(",", names));
        // il binario incorporato è identico byte per byte
        var emb = z.Entries.First(e => e.FullName.StartsWith("ppt/slides/charts/embeddings/"));
        using var ms = new MemoryStream(); using (var s = emb.Open()) s.CopyTo(ms);
        Assert.Equal(new byte[] { 0x50, 0x4B, 3, 4, 1, 2, 3, 4, 5 }, ms.ToArray());
        Assert.DoesNotContain(RoundTripTests.NewErrors(f, outp), e => !e.Contains("Chart") && !e.Contains("diagram", StringComparison.OrdinalIgnoreCase));
        // salvataggio senza modifiche: package identico
        var plain = Path.Combine(Env.Work, "chart-smartart-plain.pptx");
        using (var d = PptxDocument.Open(f)) d.SaveAs(plain);
        Assert.Empty(PackageCompare.Diff(f, plain));
        // l'export PDF non va in crash con questi oggetti
        using (var d = PptxDocument.Open(f)) PdfExporter.Export(d, Path.Combine(Env.Work, "chart-smartart.pdf"));
        Assert.True(new FileInfo(Path.Combine(Env.Work, "chart-smartart.pdf")).Length > 500);
    }

    [Fact]
    public void DuplicateSlide_WithChartAndSmartArt_CopiesTheParts_AndDeleteKeepsTheOriginal()
    {
        var f = Build(); var outp = Path.Combine(Env.Work, "chart-dup.pptx");
        using (var d = PptxDocument.Open(f))
        {
            int k = d.Slides.ToList().FindIndex(x => x.Parts.Any(p => p.OpenXmlPart is ChartPart));
            int copy = d.DuplicateSlide(k);
            var a = d.Slides[k].Parts.Where(p => p.OpenXmlPart.RelationshipType.EndsWith("/chart")).Select(p => p.OpenXmlPart).Single();
            var b = d.Slides[copy].Parts.Where(p => p.OpenXmlPart.RelationshipType.EndsWith("/chart")).Select(p => p.OpenXmlPart).Single();
            Assert.NotSame(a, b);   // part distinte
            d.DeleteSlide(k);        // l'originale se ne va, la copia conserva grafico e incorporato
            d.SaveAs(outp);
        }
        using var r = PptxDocument.Open(outp);
        var sl = r.Slides.Single(x => x.Parts.Any(p => p.OpenXmlPart.RelationshipType.EndsWith("/chart")));
        var ch = sl.Parts.Single(p => p.OpenXmlPart.RelationshipType.EndsWith("/chart")).OpenXmlPart;
        using var ms = new MemoryStream(); using (var es = ch.Parts.Single().OpenXmlPart.GetStream()) es.CopyTo(ms);
        Assert.Equal(new byte[] { 0x50, 0x4B, 3, 4, 1, 2, 3, 4, 5 }, ms.ToArray());   // incorporato copiato con la slide
        Assert.Equal(4, sl.Parts.Count(p => p.OpenXmlPart.RelationshipType.Contains("/diagram")));
    }

    [Fact]
    public void DeleteAllSlides_ThenSave_ReopensWithZeroSlides()
    {
        var f = Env.Copy("mixed-realistic-1.pptx", "-delall");
        using (var d = PptxDocument.Open(f)) { while (d.Slides.Count > 0) d.DeleteSlide(0); d.SaveAs(f); }
        using var r = PptxDocument.Open(f); Assert.Empty(r.Slides);
    }
}
