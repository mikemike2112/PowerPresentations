using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using PdfSharp.Pdf.IO;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Tests;

public class EditTests
{
    static string Out(string n) => Path.Combine(Env.Work, n + ".pptx");
    static int FirstText(PptxDocument d, int s) => new SlideRenderer(d).Render(s).Hits.First(h => h.HasText).Id;

    [Fact]
    public void Reorder_TouchesOnlySldIdLst_PartsRelationsNotesAndCommentsKept()
    {
        var src = Env.CorpusPath("mixed-realistic-1.pptx"); var dst = Out("reorder-core");
        using (var d = PptxDocument.Open(src)) { d.MoveSlide(2, 0); d.SaveAs(dst); }
        var diff = PackageCompare.Diff(src, dst);
        Assert.Equal(new[] { "XML diverso: /ppt/presentation.xml" }, diff.ToArray());
        Assert.Empty(RoundTripTests.NewErrors(src, dst));
        var p = XElement.Parse(new StreamReader(ZipFile.OpenRead(dst).GetEntry("ppt/presentation.xml")!.Open()).ReadToEnd());
        Assert.Equal(new[] { "258", "256", "257" }, p.Descendants().Where(e => e.Name.LocalName == "sldId").Select(e => (string)e.Attribute("id")!).ToArray());
    }

    [Fact]
    public void Reorder_WithCommentsAndSections_KeepsEverything()
    {
        foreach (var f in new[] { "comments.pptx", "sections.pptx", "notes-speaker.pptx" })
        {
            var src = Env.CorpusPath(f); var dst = Out("reorder-" + f);
            using (var d = PptxDocument.Open(src)) { if (d.Slides.Count > 1) d.MoveSlide(0, d.Slides.Count - 1); d.SaveAs(dst); }
            Assert.DoesNotContain(PackageCompare.Diff(src, dst), x => !x.Contains("presentation.xml"));
            Assert.Empty(RoundTripTests.NewErrors(src, dst));
        }
    }

    [Fact]
    public void DeleteSlide_RemovesNotesAndSectionReference_Valid()
    {
        var src = Env.CorpusPath("sections.pptx"); var dst = Out("del-core");
        using (var d = PptxDocument.Open(src)) { d.DeleteSlide(1); d.SaveAs(dst); }
        using var z = ZipFile.OpenRead(dst);
        Assert.Equal(3, z.Entries.Count(e => e.FullName.StartsWith("ppt/slides/slide") && e.FullName.EndsWith(".xml")));
        Assert.Empty(RoundTripTests.NewErrors(src, dst));
        var notes = Env.CorpusPath("notes-speaker.pptx"); var dst2 = Out("del-notes");
        using (var d = PptxDocument.Open(notes)) { d.DeleteSlide(0); d.SaveAs(dst2); }
        using var z2 = ZipFile.OpenRead(dst2);
        Assert.Equal(1, z2.Entries.Count(e => e.FullName.StartsWith("ppt/notesSlides/") && e.FullName.EndsWith(".xml")));
        Assert.Empty(RoundTripTests.NewErrors(notes, dst2));
    }

    [Fact]
    public void AddDuplicateAndEdit_ProduceValidFiles()
    {
        var src = Env.CorpusPath("mixed-realistic-1.pptx"); var dst = Out("add-dup");
        using (var d = PptxDocument.Open(src))
        {
            var n = d.AddSlide(d.Layouts[1], 0); d.SetShapeText(n, FirstText(d, n), "Titolo nuovo\nsecondo paragrafo");
            d.DuplicateSlide(2); d.SetNotes(n, "una nota"); d.SetTransition(n, "wipe", "r", "slow");
            d.AddTextBox(0, 10, 10, 100, 30, "T"); d.AddShape(0, "roundRect", 10, 50, 100, 60, "S");
            d.AddPicture(0, ComEditFixture.Png, 200, 50, 40, 40); d.AddTable(0, 3, 3, 100, 300, 400);
            d.SaveAs(dst);
        }
        Assert.Empty(RoundTripTests.NewErrors(src, dst));
        using var r = PptxDocument.Open(dst);
        Assert.Equal(5, r.Slides.Count);
        Assert.Equal("wipe", r.GetTransition(1).Kind); Assert.Equal("r", r.GetTransition(1).Dir); Assert.Equal("una nota", r.GetNotes(1));
        Assert.Equal("Titolo nuovo\nsecondo paragrafo", r.GetShapeText(1, FirstText(r, 1)));
    }

    [Fact]
    public void Duplicate_CopiesUnsavedEditsOfTheSource()
    {
        using var d = PptxDocument.Open(Env.CorpusPath("title-slide.pptx"));
        d.SetTransition(0, "fade"); d.SetShapeText(0, FirstText(d, 0), "Modificato");
        var i = d.DuplicateSlide(0);
        Assert.Equal("fade", d.GetTransition(i).Kind); Assert.Equal("Modificato", d.GetShapeText(i, FirstText(d, i)));
    }

    [Fact]
    public void SetShapeText_KeepsRunFormatAndParagraphProperties()
    {
        var dst = Out("fmt");
        using (var d = PptxDocument.Open(Env.CorpusPath("text-formatting.pptx")))
        {
            var id = FirstText(d, 0);
            d.SetShapeText(0, id, "Nuovo 1\nNuovo 2");
            d.SetFormat(0, id, bold: true, sizePt: 30, colorHex: "FF0000");
            d.SaveAs(dst);
        }
        using var z = ZipFile.OpenRead(dst);
        var x = new StreamReader(z.GetEntry("ppt/slides/slide1.xml")!.Open()).ReadToEnd();
        Assert.Contains("Nuovo 1", x); Assert.Contains("sz=\"3000\"", x); Assert.Contains("FF0000", x);
        Assert.Empty(RoundTripTests.NewErrors(Env.CorpusPath("text-formatting.pptx"), dst));
    }

    [Fact]
    public void SetShapeText_OnAutofitPlaceholder_StoresFontScale()
    {
        var dst = Out("autofit-core");
        using (var d = PptxDocument.Open(Env.CorpusPath("autofit-shrink.pptx")))
        {
            var id = new SlideRenderer(d).Render(0).Hits.Last(h => h.HasText).Id;
            d.SetShapeText(0, id, string.Join("\n", Enumerable.Range(1, 20).Select(i => "Riga numero " + i + " con testo piuttosto lungo per riempire lo spazio")));
            d.SaveAs(dst);
        }
        using var z = ZipFile.OpenRead(dst);
        var x = new StreamReader(z.GetEntry("ppt/slides/slide1.xml")!.Open()).ReadToEnd();
        Assert.Matches("normAutofit fontScale=\"\\d+\"", x);
        Assert.Empty(RoundTripTests.NewErrors(Env.CorpusPath("autofit-shrink.pptx"), dst));
    }

    [Fact]
    public void New_FromTemplate_IsValidAndSavable()
    {
        var dst = Out("new-core");
        using (var d = PptxDocument.New()) { Assert.Single(d.Slides); Assert.True(d.Layouts.Count == 6); d.SaveAs(dst); }
        using var r = PptxDocument.Open(dst);
        Assert.Empty(RoundTripTests.Errors(dst));
        using var z = ZipFile.OpenRead(dst);
        foreach (var e in z.Entries) Assert.DoesNotContain("nomeutente", new StreamReader(e.Open()).ReadToEnd());   // nessun dato personale; modello generato da codice, non da file Office
    }

    [Fact]
    public void Save_IsAtomic_FailureLeavesOriginalIntactAndNoTempFile()
    {
        var f = Env.Copy("title-slide.pptx", "-atomic"); var before = File.ReadAllBytes(f);
        using var d = PptxDocument.Open(f);
        d.SetShapeText(0, FirstText(d, 0), "cambiato");
        using (new FileStream(f, FileMode.Open, FileAccess.Read, FileShare.None))   // destinazione bloccata: la sostituzione deve fallire
            Assert.ThrowsAny<IOException>(() => d.SaveAs(f));
        Assert.Equal(before, File.ReadAllBytes(f));
        Assert.Empty(Directory.GetFiles(Env.Work, "title-slide-atomic.pptx.tmp*"));
        d.SaveAs(f);   // sbloccato: ora riesce
        using var r = PptxDocument.Open(f); Assert.Equal("cambiato", r.GetShapeText(0, FirstText(r, 0)));
    }

    [Fact]
    public void MalformedInput_GivesClearMessage_NoCrash()
    {
        var garbage = Path.Combine(Env.Work, "garbage.pptx"); File.WriteAllText(garbage, "questo non è uno zip");
        Assert.Throws<PptxFormatException>(() => PptxDocument.Open(garbage));
        var emptyZip = Path.Combine(Env.Work, "empty.pptx"); using (ZipFile.Open(emptyZip, ZipArchiveMode.Create)) { }
        Assert.Throws<PptxFormatException>(() => PptxDocument.Open(emptyZip));
        var trunc = Path.Combine(Env.Work, "trunc.pptx"); var b = File.ReadAllBytes(Env.CorpusPath("mixed-realistic-1.pptx")); File.WriteAllBytes(trunc, b[..(b.Length / 2)]);
        Assert.Throws<PptxFormatException>(() => PptxDocument.Open(trunc));
        var badXml = Path.Combine(Env.Work, "badxml.pptx"); File.Copy(Env.CorpusPath("title-slide.pptx"), badXml, true);
        using (var z = ZipFile.Open(badXml, ZipArchiveMode.Update)) { z.GetEntry("ppt/presentation.xml")!.Delete(); using var w = new StreamWriter(z.CreateEntry("ppt/presentation.xml").Open()); w.Write("<p:presentation><oops"); }
        Assert.Throws<PptxFormatException>(() => PptxDocument.Open(badXml));
        Assert.Throws<PptxFormatException>(() => PptxDocument.Open(Path.Combine(Env.Work, "non-esiste.pptx")));
    }

    [Fact]
    public void MalformedSlideXml_DoesNotCrashRenderer()
    {
        var f = Path.Combine(Env.Work, "badslide.pptx"); File.Copy(Env.CorpusPath("title-slide.pptx"), f, true);
        using (var z = ZipFile.Open(f, ZipArchiveMode.Update))
        {
            var e = z.GetEntry("ppt/slides/slide1.xml")!; string t; using (var r = new StreamReader(e.Open())) t = r.ReadToEnd(); e.Delete();
            using var w = new StreamWriter(z.CreateEntry("ppt/slides/slide1.xml").Open()); w.Write(t.Replace("sz=\"6000\"", "sz=\"abc\"").Replace("<a:off x=\"", "<a:off x=\"zz"));
        }
        using var d = PptxDocument.Open(f);
        var dl = new SlideRenderer(d).Render(0);           // nessuna eccezione: le forme malformate sono ignorate e segnalate
        Assert.NotEmpty(dl.Warnings.Concat(dl.Items.Select(_ => "x")));
        Assert.Contains(dl.Items, i => i is TextItem t && t.Text == "Sottotitolo di prova");   // il resto della slide si disegna
    }
}

public class PdfTests
{
    [Theory]
    [InlineData("mixed-realistic-1.pptx", 3)]
    [InlineData("hidden-slide.pptx", 2)]   // le slide nascoste non si esportano
    [InlineData("layouts-all.pptx", 11)]
    [InlineData("size-4x3.pptx", 1)]
    public void Export_ProducesPdfWithExpectedPages(string name, int pages)
    {
        var pdf = Path.Combine(Env.Work, Path.GetFileNameWithoutExtension(name) + ".pdf");
        using (var d = PptxDocument.Open(Env.CorpusPath(name))) PdfExporter.Export(d, pdf);
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(File.ReadAllBytes(pdf), 0, 4));
        using var doc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import);
        Assert.Equal(pages, doc.PageCount);
        if (name == "size-4x3.pptx") { Assert.Equal(720, doc.Pages[0].Width.Point, 0.5); Assert.Equal(540, doc.Pages[0].Height.Point, 0.5); }
    }

    [Fact]
    public void Export_AllCorpus_NoException()
    {
        foreach (var f in Directory.GetFiles(Env.Corpus, "*.pptx"))
        {
            var pdf = Path.Combine(Env.Work, "all-" + Path.GetFileNameWithoutExtension(f) + ".pdf");
            using var d = PptxDocument.Open(f); PdfExporter.Export(d, pdf);
            Assert.True(new FileInfo(pdf).Length > 500, f);
        }
    }
}

public class SettingsAndUndoTests
{
    [Fact]
    public void Settings_RoundTrip_AndCorruptedFilesFallBackToDefaults()
    {
        var f = Path.Combine(Env.Work, "settings.json");
        Assert.True(Core.AppSettings.Load(f).Dark);                         // nessun file: scuro
        new Core.AppSettings { Dark = false }.Save(f);
        Assert.False(Core.AppSettings.Load(f).Dark);
        foreach (var bad in new[] { "{ non json", "", "[1,2,3]", "{\"Dark\":\"forse\"}", new string('x', 5000) })
        {
            File.WriteAllText(f, bad);
            Assert.True(Core.AppSettings.Load(f).Dark, bad);                // mai un'eccezione
        }
        File.Delete(f); Directory.CreateDirectory(f);                       // percorso che è una cartella: non si può leggere né scrivere
        Assert.True(Core.AppSettings.Load(f).Dark); new Core.AppSettings().Save(f);
    }

    [Fact]
    public void SnapshotRestore_UndoesAndRedoesEdits()
    {
        using var d = PptxDocument.Open(Env.CorpusPath("textbox-basic.pptx"));
        var id = new SlideRenderer(d).Render(0).Hits.First(h => h.HasText).Id;
        var before = d.GetShapeText(0, id); var s0 = d.Snapshot();
        d.SetShapeText(0, id, "cambiato"); d.AddSlide(d.Layouts[1], 0); var s1 = d.Snapshot();
        Assert.Equal(2, d.Slides.Count);
        d.Restore(s0); Assert.Single(d.Slides); Assert.Equal(before, d.GetShapeText(0, id));
        d.Restore(s1); Assert.Equal(2, d.Slides.Count); Assert.Equal("cambiato", d.GetShapeText(0, id));
    }

    [Fact]
    public void TableOptions_AlignmentAndCellText_AreWritten()
    {
        var dst = Path.Combine(Env.Work, "table-ui.pptx");
        using (var d = PptxDocument.Open(Env.CorpusPath("table-basic.pptx")))
        {
            var r = new SlideRenderer(d).Render(0); var frame = r.TableCells.First().FrameId;
            Assert.True(d.GetTableFlag(0, frame, "firstRow"));
            d.SetTableFlag(0, frame, "firstRow", false); d.SetTableFlag(0, frame, "lastCol", true);
            d.SetCellText(0, frame, 1, 1, "X"); Assert.Equal("X", d.GetCellText(0, frame, 1, 1));
            var title = r.Hits.First(h => h.Id != frame && h.HasText).Id; d.SetAlignment(0, title, "r");
            d.SaveAs(dst);
        }
        using var z = System.IO.Compression.ZipFile.OpenRead(dst);
        var x = new StreamReader(z.GetEntry("ppt/slides/slide1.xml")!.Open()).ReadToEnd();
        Assert.Contains("lastCol=\"1\"", x); Assert.Contains("firstRow=\"0\"", x); Assert.Contains("algn=\"r\"", x);
        Assert.Empty(RoundTripTests.NewErrors(Env.CorpusPath("table-basic.pptx"), dst));
    }
}

public class TransitionDataLossTests
{
    static string TransitionXml(string file, int slide)
    {
        using var z = System.IO.Compression.ZipFile.OpenRead(file);
        var x = new StreamReader(z.GetEntry($"ppt/slides/slide{slide}.xml")!.Open()).ReadToEnd();
        var m = System.Text.RegularExpressions.Regex.Match(x, "<mc:AlternateContent[^>]*>(?:(?!</mc:AlternateContent>).)*transition.*?</mc:AlternateContent>|<p:transition.*?</p:transition>", System.Text.RegularExpressions.RegexOptions.Singleline);
        return m.Value.Replace(" />", "/>");   // l'SDK serializza gli elementi vuoti con uno spazio: stessa semantica
    }

    [Fact]
    public void UnchangedTransition_StaysByteIdentical_AfterOtherEdits()
    {
        var src = Env.CorpusPath("transitions.pptx"); var dst = Path.Combine(Env.Work, "tr-keep.pptx");
        using (var d = PptxDocument.Open(src))
        {
            d.SetShapeText(1, new SlideRenderer(d).Render(1).Hits.First(h => h.HasText).Id, "altro testo"); d.SetNotes(2, "nota"); d.MoveSlide(5, 0); d.SaveAs(dst);
        }
        // la transizione non toccata resta identica (con p14:dur)
        var a = TransitionXml(src, 2); Assert.Contains("p14:dur", a);
        Assert.Equal(a, TransitionXml(dst, 2));
        Assert.Equal(TransitionXml(src, 3), TransitionXml(dst, 3));
    }

    [Fact]
    public void ChangedTransition_KeepsDurationInAlternateContent_AndIsValid()
    {
        var src = Env.CorpusPath("transitions.pptx"); var dst = Path.Combine(Env.Work, "tr-change.pptx");
        using (var d = PptxDocument.Open(src)) { Assert.Equal("cover", d.GetTransition(1).Kind); d.SetTransition(1, "cover", "r", "slow"); d.SaveAs(dst); }
        var x = TransitionXml(dst, 2);
        Assert.Contains("<mc:Fallback>", x); Assert.Contains("p14:dur=\"2000\"", x); Assert.Contains("dir=\"r\"", x);
        Assert.Empty(RoundTripTests.NewErrors(src, dst));
        using var r = PptxDocument.Open(dst); Assert.Equal("r", r.GetTransition(1).Dir);
    }

    [Fact]
    public void ChangingToNone_RemovesTheWholeAlternateContent()
    {
        var dst = Path.Combine(Env.Work, "tr-none.pptx");
        using (var d = PptxDocument.Open(Env.CorpusPath("transitions.pptx"))) { d.SetTransition(1, "none"); d.SaveAs(dst); }
        Assert.Equal("", TransitionXml(dst, 2));
    }
}

public class CommentsTests
{
    [Fact]
    public void DuplicateSlide_LeavesOriginalCommentsIntact_CopyHasNone_FileValid()
    {
        var src = Env.CorpusPath("comments.pptx"); var dst = Path.Combine(Env.Work, "comments-dup.pptx");
        using (var d = PptxDocument.Open(src)) { d.DuplicateSlide(0); d.SaveAs(dst); }
        using var z = System.IO.Compression.ZipFile.OpenRead(dst);
        Assert.Single(z.Entries, e => e.FullName.StartsWith("ppt/comments/"));                   // il commento dell'originale c'è ancora, uno solo
        var rels1 = new StreamReader(z.GetEntry("ppt/slides/_rels/slide1.xml.rels")!.Open()).ReadToEnd();
        Assert.Contains("relationships/comments", rels1);
        Assert.DoesNotContain("comments", new StreamReader(z.Entries.First(e => e.FullName.StartsWith("ppt/slides/_rels/") && e.FullName != "ppt/slides/_rels/slide1.xml.rels").Open()).ReadToEnd());
        Assert.Empty(RoundTripTests.NewErrors(src, dst));
    }
}
