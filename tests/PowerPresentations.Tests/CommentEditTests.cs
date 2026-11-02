using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using PowerPresentations.Core;

namespace PowerPresentations.Tests;

/// <summary>Modifica ed eliminazione dei commenti, moderni (con risposte) e classici.</summary>
public class CommentEditTests
{
    static readonly XNamespace P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
    const string PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    static string Out(string n) => Path.Combine(Env.Work, "cedit-" + n + ".pptx");

    static void Feed(OpenXmlPart p, string xml) { using var s = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml)); p.FeedData(s); }

    static List<string> Entries(string f) { using var z = ZipFile.OpenRead(f); return z.Entries.Select(e => e.FullName).ToList(); }
    static string Entry(string f, string name) { using var z = ZipFile.OpenRead(f); return new StreamReader(z.GetEntry(name)!.Open()).ReadToEnd(); }
    static bool HasCommentParts(string f) => Entries(f).Any(e => e.StartsWith("ppt/comments/"));

    /// <summary>Salva, valida (nessun errore nuovo rispetto a src se dato) e riapre.</summary>
    static PptxDocument Save(PptxDocument d, string name, string? orig = null)
    {
        var f = Out(name); d.SaveAs(f);
        var errs = orig == null ? RoundTripTests.Errors(f) : RoundTripTests.NewErrors(orig, f);
        Assert.True(errs.Count == 0, string.Join(" ## ", errs));
        return PptxDocument.Open(f);
    }

    static void AssertNoOrphans(string f)
    {
        // ogni relazione di slide verso commenti punta a una part esistente; ogni authorId usato esiste negli autori
        var names = Entries(f);
        foreach (var r in names.Where(n => Path.GetDirectoryName(n)!.Replace('\\', '/') == "ppt/slides/_rels"))
            foreach (var m in System.Text.RegularExpressions.Regex.Matches(Entry(f, r), "Target=\"\\.\\./(comments/[^\"]+)\"").Cast<System.Text.RegularExpressions.Match>()) Assert.Contains("ppt/" + m.Groups[1].Value, names);
        var authors = names.Contains("ppt/authors.xml") ? XElement.Parse(Entry(f, "ppt/authors.xml")).Elements(P188 + "author").Select(a => (string)a.Attribute("id")!).ToHashSet() : new();
        foreach (var c in names.Where(n => n.StartsWith("ppt/comments/modern")))
            foreach (var e in XElement.Parse(Entry(f, c)).Descendants().Where(e => e.Name == P188 + "cm" || e.Name == P188 + "reply")) Assert.Contains((string)e.Attribute("authorId")!, authors);
    }

    static void AddReply(PptxDocument d, int slide, string cmIndex, string text)
    {
        var cp = d.Slides[slide].GetPartsOfType<PowerPointCommentPart>().Single(); var x = PptxDocument.Xml(cp);
        var cm = x.Elements(P188 + "cm").ElementAt(int.Parse(cmIndex));
        var lst = cm.Element(P188 + "replyLst") ?? new XElement(P188 + "replyLst"); if (lst.Parent == null) cm.Element(P188 + "pos")!.AddAfterSelf(lst);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        lst.Add(new XElement(P188 + "reply", new XAttribute("id", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"), new XAttribute("authorId", (string)cm.Attribute("authorId")!), new XAttribute("created", "2026-10-02T08:00:00.000"),
            new XElement(P188 + "txBody", new XElement(a + "bodyPr"), new XElement(a + "lstStyle"), new XElement(a + "p", new XElement(a + "r", new XElement(a + "rPr", new XAttribute("lang", "it-IT")), new XElement(a + "t", text))))));
        Feed(cp, x.ToString(SaveOptions.DisableFormatting));
    }

    static void AddLegacy(PptxDocument d, int slide)
    {
        Feed(d.PresentationPart.AddNewPart<CommentAuthorsPart>(), $"<p:cmAuthorLst xmlns:p=\"{PNs}\"><p:cmAuthor id=\"0\" name=\"Autore\" initials=\"A\" lastIdx=\"3\" clrIdx=\"0\"/></p:cmAuthorLst>");
        const string ext = "<p:extLst><p:ext uri=\"{C676402C-5697-4E1C-873F-D02D1690AC5C}\"><p15:threadingInfo xmlns:p15=\"http://schemas.microsoft.com/office/powerpoint/2012/main\" timeZoneBias=\"-120\"><p15:parentCm authorId=\"0\" idx=\"1\"/></p15:threadingInfo></p:ext></p:extLst>";
        Feed(d.Slides[slide].AddNewPart<SlideCommentsPart>(), $"<p:cmLst xmlns:p=\"{PNs}\">"
            + "<p:cm authorId=\"0\" dt=\"2026-10-01T10:00:00.000\" idx=\"1\"><p:pos x=\"10\" y=\"10\"/><p:text>Classico uno</p:text></p:cm>"
            + $"<p:cm authorId=\"0\" dt=\"2026-10-01T11:00:00.000\" idx=\"2\"><p:pos x=\"10\" y=\"10\"/><p:text>Risposta classica</p:text>{ext}</p:cm>"
            + "<p:cm authorId=\"0\" dt=\"2026-10-01T12:00:00.000\" idx=\"3\"><p:pos x=\"20\" y=\"20\"/><p:text>Classico tre</p:text></p:cm></p:cmLst>");
    }

    [Fact]
    public void Modern_AppCreated_EditKeepsAuthorDateAndOthers_DeleteLastRemovesPartAndRelation()
    {
        using var d = PptxDocument.New();
        d.AddComment(0, "uno"); d.AddComment(0, "due\nsu righe"); d.AddComment(0, "tre");
        var before = d.GetComments(0);
        d.EditComment(0, before[1].Id, "due cambiato\nancora");
        var after = d.GetComments(0);
        Assert.Equal("due cambiato\nancora", after[1].Text);
        Assert.Equal(before[1] with { Text = after[1].Text }, after[1]);                                  // autore, data, id invariati
        Assert.Equal(before[0], after[0]); Assert.Equal(before[2], after[2]);                              // gli altri non si toccano
        using (var r = Save(d, "m-edit")) Assert.Equal("due cambiato\nancora", r.GetComments(0)[1].Text);
        Assert.Throws<ArgumentException>(() => d.EditComment(0, before[0].Id, "  "));
        Assert.Throws<KeyNotFoundException>(() => d.EditComment(0, "{nope}", "x"));
        Assert.Throws<KeyNotFoundException>(() => d.DeleteComment(0, "{nope}"));

        d.DeleteComment(0, before[0].Id);
        Assert.Equal(new[] { "due cambiato\nancora", "tre" }, d.GetComments(0).Select(c => c.Text));
        using (var r = Save(d, "m-del1")) Assert.Equal(2, r.GetComments(0).Count);
        d.DeleteComment(0, before[1].Id); d.DeleteComment(0, before[2].Id);
        Assert.Empty(d.GetComments(0));
        var f = Out("m-del-all"); using (var r = Save(d, "m-del-all")) Assert.Empty(r.GetComments(0));
        Assert.False(HasCommentParts(f));
        Assert.DoesNotContain("comments", Entry(f, "ppt/slides/_rels/slide1.xml.rels"));
        Assert.DoesNotContain("commentRel", Entry(f, "ppt/slides/slide1.xml"));
        Assert.DoesNotContain("ms-powerpoint.comments", Entry(f, "[Content_Types].xml"));
        AssertNoOrphans(f);
        d.AddComment(0, "di nuovo"); using (var r = Save(d, "m-readd")) Assert.Single(r.GetComments(0));   // dopo l'ultima eliminazione si può commentare ancora
    }

    [Fact]
    public void Modern_OfficeMadeCorpus_EditDelete_ValidAndAuthorKept()
    {
        var src = Env.CorpusPath("comments.pptx");
        using var d = PptxDocument.Open(src); var c = d.GetComments(0); Assert.Single(c); Assert.Equal("Corpus", c[0].Author); Assert.Equal("Commento di prova", c[0].Text);
        d.EditComment(0, c[0].Id, "Testo nuovo");
        using (var r = Save(d, "corpus-edit", src)) { var e = r.GetComments(0)[0]; Assert.Equal(c[0] with { Text = "Testo nuovo" }, e); }
        d.DeleteComment(0, c[0].Id);
        var f = Out("corpus-del"); using (var r = Save(d, "corpus-del", src)) Assert.Empty(r.GetComments(0));
        Assert.False(HasCommentParts(f)); AssertNoOrphans(f);
        Assert.Contains("Corpus", Entry(f, "ppt/authors.xml"));   // gli autori restano coerenti (PowerPoint li conserva)
    }

    [Fact]
    public void Modern_Replies_ListedEditedAndDeletedWithParent()
    {
        using var d = PptxDocument.New();
        d.AddComment(0, "padre"); d.AddComment(0, "altro");
        AddReply(d, 0, "0", "risposta 1"); AddReply(d, 0, "0", "risposta 2");
        var c = d.GetComments(0);
        Assert.Equal(new[] { "padre", "risposta 1", "risposta 2", "altro" }, c.Select(x => x.Text));
        Assert.Equal(2, c[0].Replies); Assert.True(c[1].IsReply && c[2].IsReply && !c[3].IsReply);
        d.EditComment(0, c[1].Id, "risposta modificata");
        Assert.Equal(c[1] with { Text = "risposta modificata" }, d.GetComments(0)[1]);
        using (var r = Save(d, "rep-edit")) Assert.Equal(4, r.GetComments(0).Count);
        d.DeleteComment(0, c[2].Id);                                                   // solo una risposta
        Assert.Equal(new[] { "padre", "risposta modificata", "altro" }, d.GetComments(0).Select(x => x.Text));
        d.DeleteComment(0, c[0].Id);                                                   // il padre porta via le risposte
        Assert.Equal(new[] { "altro" }, d.GetComments(0).Select(x => x.Text));
        var f = Out("rep-del"); using (var r = Save(d, "rep-del")) Assert.Single(r.GetComments(0));
        Assert.DoesNotContain("risposta", Entry(f, Entries(f).Single(e => e.StartsWith("ppt/comments/")))); AssertNoOrphans(f);
    }

    [Fact]
    public void Legacy_Listed_Edited_Deleted_WithRepliesAndLastRemovesPart()
    {
        using var d = PptxDocument.FromBytes(File.ReadAllBytes(Env.CorpusPath("title-slide.pptx")));
        AddLegacy(d, 0);
        var c = d.GetComments(0);
        Assert.Equal(new[] { "Classico uno", "Risposta classica", "Classico tre" }, c.Select(x => x.Text));
        Assert.Equal("Autore", c[0].Author); Assert.Equal(1, c[0].Replies); Assert.True(c[1].IsReply);
        d.EditComment(0, c[2].Id, "Tre cambiato");
        Assert.Equal(c[2] with { Text = "Tre cambiato" }, d.GetComments(0)[2]); Assert.Equal(c[0], d.GetComments(0)[0]);
        using (var r = Save(d, "leg-edit")) Assert.Equal("Tre cambiato", r.GetComments(0)[2].Text);
        d.DeleteComment(0, c[0].Id);                                                   // via anche la risposta classica
        Assert.Equal(new[] { "Tre cambiato" }, d.GetComments(0).Select(x => x.Text));
        d.AddComment(0, "nuovo classico");                                             // sulla slide con commenti classici se ne aggiunge uno classico
        Assert.Equal(2, d.GetComments(0).Count);
        d.DeleteComment(0, d.GetComments(0)[1].Id); d.DeleteComment(0, c[2].Id);
        var f = Out("leg-del"); using (var r = Save(d, "leg-del")) Assert.Empty(r.GetComments(0));
        Assert.DoesNotContain("comments/", Entry(f, "ppt/slides/_rels/slide1.xml.rels"));
        Assert.DoesNotContain(Entries(f), e => e.StartsWith("ppt/comments/")); Assert.Contains("commentAuthors", string.Join(",", Entries(f)));
    }

    [Fact]
    public void SnapshotRestore_BringsBackDeletedAndEditedComments()
    {
        using var d = PptxDocument.New();
        d.AddComment(0, "originale"); var id = d.GetComments(0)[0].Id;
        var snap = d.Snapshot(); d.EditComment(0, id, "cambiato");
        var snap2 = d.Snapshot(); d.DeleteComment(0, id); Assert.Empty(d.GetComments(0));
        d.Restore(snap2); Assert.Equal("cambiato", d.GetComments(0)[0].Text);
        d.Restore(snap); Assert.Equal("originale", d.GetComments(0)[0].Text);
    }

    [Fact]
    public void Slides_DuplicateDeletePaste_StayConsistentWithComments()
    {
        var src = Env.CorpusPath("comments.pptx");
        using var d = PptxDocument.Open(src);
        int dup = d.DuplicateSlide(0);
        Assert.Single(d.GetComments(0)); Assert.Empty(d.GetComments(dup));                 // la copia non porta commenti, l'originale li tiene
        d.AddComment(dup, "sulla copia");
        d.PasteShapes(dup, d.CopyShapes(0, new[] { d.ShapeIds(0)[0] }));                    // incolla forme: i commenti non si muovono
        Assert.Single(d.GetComments(0)); Assert.Single(d.GetComments(dup));
        var f = Out("slides-dup"); using (var r = Save(d, "slides-dup", src)) Assert.Equal(2, r.Slides.Count);
        AssertNoOrphans(f); Assert.Equal(2, Entries(f).Count(e => e.StartsWith("ppt/comments/")));
        d.DeleteSlide(0);                                                                  // elimina la slide col commento: sparisce anche la sua part
        var f2 = Out("slides-del"); using (var r = Save(d, "slides-del", src)) Assert.Single(r.Slides);
        Assert.Single(Entries(f2), e => e.StartsWith("ppt/comments/")); AssertNoOrphans(f2);
        d.DeleteSlide(0); d.AddSlide(d.Layouts[0], -1);
        Assert.False(HasCommentParts(Save(d, "slides-del-all", src) is var q ? Out("slides-del-all") : ""));
    }
}
