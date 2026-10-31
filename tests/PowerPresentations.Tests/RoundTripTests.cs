using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Tests;

public class RoundTripTests
{
    public static List<string> Errors(string path)
    {
        using var d = PresentationDocument.Open(path, false);
        return new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(d).Select(e => $"{e.ErrorType}|{e.Description}|{e.Path?.XPath}").ToList();
    }

    public static List<string> NewErrors(string original, string saved)
    {
        var before = Errors(original); var after = Errors(saved);
        return after.GroupBy(x => x).Where(g => g.Count() > before.Count(b => b == g.Key)).Select(g => g.Key).ToList();
    }

    [Theory, MemberData(nameof(Env.CorpusFiles), MemberType = typeof(Env))]
    public void OpenSave_NoChanges_ValidatesAndPreservesEverything(string name)
    {
        var src = Env.CorpusPath(name); var dst = Path.Combine(Env.Work, "rt-" + name);
        using (var d = PptxDocument.Open(src)) d.SaveAs(dst);
        // a. validatore: nessun errore nuovo rispetto all'originale (quelli dell'originale sono elencati in COMPAT.md)
        var added = NewErrors(src, dst);
        Assert.True(added.Count == 0, $"{name}: errori nuovi del validatore:\n" + string.Join("\n", added));
        // d. stessa lista di part/relazioni, XML canonicalizzato identico
        var diff = PackageCompare.Diff(src, dst);
        Assert.True(diff.Count == 0, $"{name}: " + string.Join("; ", diff));
    }

    [Fact]
    public void OriginalValidatorErrors_AreDumpedForCompat()
    {
        var lines = new List<string>();
        foreach (var f in Directory.GetFiles(Env.Corpus, "*.pptx").OrderBy(f => f))
            lines.AddRange(Errors(f).Select(e => Path.GetFileName(f) + " :: " + e));
        File.WriteAllLines(Path.Combine(Env.Work, "validator-baseline.txt"), lines);
    }

    [Theory, MemberData(nameof(Env.CorpusFiles), MemberType = typeof(Env))]
    public void Corpus_Metadata_AuthorIsCorpusAndNoPersonalData(string name)
    {
        using var z = ZipFile.OpenRead(Env.CorpusPath(name));
        var core = XElement.Parse(new StreamReader(z.GetEntry("docProps/core.xml")!.Open()).ReadToEnd());
        Assert.Equal("Corpus", core.Descendants().First(e => e.Name.LocalName == "creator").Value);
        Assert.Equal("Corpus", core.Descendants().First(e => e.Name.LocalName == "lastModifiedBy").Value);
        foreach (var e in z.Entries.Where(e => e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels")))
        {
            var t = new StreamReader(e.Open()).ReadToEnd();
            Assert.DoesNotContain("nomeutente", t, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("@gmail", t, StringComparison.OrdinalIgnoreCase);
        }
    }

    const string Zz = "urn:powerpresentations:test";

    /// <summary>File con part custom, elemento di namespace estraneo in mc:Ignorable, extLst e attributo sconosciuto.</summary>
    public static string BuildUnknownContentFile()
    {
        var path = Path.Combine(Env.Work, "unknown-content.pptx");
        File.Copy(Env.CorpusPath("mixed-realistic-1.pptx"), path, true);
        using var z = ZipFile.Open(path, ZipArchiveMode.Update);
        string Read(string n) { using var r = new StreamReader(z.GetEntry(n)!.Open(), new UTF8Encoding(false)); return r.ReadToEnd(); }
        void Write(string n, string t) { z.GetEntry(n)?.Delete(); using var w = new StreamWriter(z.CreateEntry(n).Open(), new UTF8Encoding(false)); w.Write(t); }
        var s = Read("ppt/slides/slide2.xml");
        s = s.Replace("<p:sld ", $"<p:sld xmlns:zz=\"{Zz}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"zz\" ");
        s = s.Replace("</p:spTree>", "<zz:marker zz:id=\"42\">dato estraneo</zz:marker></p:spTree>");
        s = s.Replace("<p:cNvPr id=\"2\"", "<p:cNvPr zz:attr=\"x\" id=\"2\"");
        s = s.Replace("</p:sld>", "<p:extLst><p:ext uri=\"{11111111-2222-3333-4444-555555555555}\"><zz:ext v=\"1\"/></p:ext></p:extLst></p:sld>");
        Write("ppt/slides/slide2.xml", s);
        Write("customData/data1.xml", $"<root xmlns=\"{Zz}\"><v>custom</v></root>");
        Write("[Content_Types].xml", Read("[Content_Types].xml").Replace("</Types>", "<Override PartName=\"/customData/data1.xml\" ContentType=\"application/vnd.test+xml\"/></Types>"));
        Write("ppt/_rels/presentation.xml.rels", Read("ppt/_rels/presentation.xml.rels").Replace("</Relationships>", "<Relationship Id=\"rIdZz\" Type=\"urn:test/custom\" Target=\"../customData/data1.xml\"/></Relationships>"));
        return path;
    }

    [Fact]
    public void UnknownContent_SurvivesEditAndSave()
    {
        var f = BuildUnknownContentFile();
        var outp = Path.Combine(Env.Work, "unknown-content-out.pptx");
        using (var d = PptxDocument.Open(f))
        {
            var id = new SlideRenderer(d).Render(1).Hits.First(h => h.HasText).Id;
            d.SetShapeText(1, id, "Testo modificato");
            d.MoveSlide(1, 0);
            d.SetNotes(0, "nota nuova");
            d.SaveAs(outp);
        }
        using var z = ZipFile.OpenRead(outp);
        var slide = new StreamReader(z.GetEntry("ppt/slides/slide2.xml")!.Open()).ReadToEnd();
        Assert.Contains("zz:marker", slide); Assert.Contains("dato estraneo", slide);
        Assert.Contains("zz:attr=\"x\"", slide);
        Assert.Contains("extLst", slide); Assert.Contains("11111111-2222-3333-4444-555555555555", slide);
        Assert.Matches("mc:Ignorable=\"[^\"]*zz", slide);
        Assert.Contains("Testo modificato", slide);
        Assert.NotNull(z.GetEntry("customData/data1.xml"));
        Assert.Contains("rIdZz", new StreamReader(z.GetEntry("ppt/_rels/presentation.xml.rels")!.Open()).ReadToEnd());
        Assert.Empty(NewErrors(f, outp));
    }
}
