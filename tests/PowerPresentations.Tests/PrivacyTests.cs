using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using PowerPresentations.Core;

namespace PowerPresentations.Tests;

public class PrivacyTests
{
    const string Name1 = "Mario Rossi", Name2 = "Anna Verdi", Comp = "Acme Segreta SRL", Mgr = "Capo Reparto", Prop = "ValoreRiservato123", PathUser = @"C:\Users\mrossi\Desktop\modello.potx";

    /// <summary>File con dati personali noti, costruito dal file del vero PowerPoint con commenti.</summary>
    public static string BuildDirty()
    {
        var path = Path.Combine(Env.Work, "privacy-dirty.pptx");
        File.Copy(Env.CorpusPath("comments.pptx"), path, true);
        using (var doc = PresentationDocument.Open(path, true))
        {
            doc.PackageProperties.Creator = Name1; doc.PackageProperties.LastModifiedBy = Name2;
            var ext = doc.ExtendedFilePropertiesPart!;
            System.Xml.Linq.XDocument x; using (var rs = ext.GetStream()) x = System.Xml.Linq.XDocument.Load(rs);
            var ns = x.Root!.GetDefaultNamespace();
            // gli elementi esistono già (vuoti) nel file di PowerPoint: si valorizzano, senza duplicarli (sarebbe un file non valido)
            foreach (var (n, v) in new[] { ("Template", PathUser), ("Manager", Mgr), ("Company", Comp) })
            {
                var el = x.Root.Element(ns + n);
                if (el != null) el.Value = v; else x.Root.Element(ns + "Application")!.AddAfterSelf(new System.Xml.Linq.XElement(ns + n, v));
            }
            using (var s = ext.GetStream(FileMode.Create)) x.Save(s);
            var cp = doc.AddCustomFilePropertiesPart();
            using (var s = cp.GetStream(FileMode.Create)) s.Write(Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}\" pid=\"2\" name=\"Cliente\"><vt:lpwstr>{Prop}</vt:lpwstr></property></Properties>"));
        }
        using (var z = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var e = z.GetEntry("ppt/authors.xml")!; string t; using (var r = new StreamReader(e.Open())) t = r.ReadToEnd(); e.Delete();
            using var w = new StreamWriter(z.CreateEntry("ppt/authors.xml").Open(), new UTF8Encoding(false)); w.Write(t.Replace("Corpus", Name1).Replace("Sorpus", "mrossi@example.org"));
        }
        return path;
    }

    static string AllText(byte[] data)
    {
        var sb = new StringBuilder();
        using var z = new ZipArchive(new MemoryStream(data), ZipArchiveMode.Read);
        foreach (var e in z.Entries) { using var s = e.Open(); using var m = new MemoryStream(); s.CopyTo(m); sb.Append(e.FullName).Append('\n').Append(Encoding.UTF8.GetString(m.ToArray())).Append('\n'); }
        return sb.ToString();
    }

    static string Sha(byte[] b) => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(b));

    [Fact]
    public void Scan_FindsKnownValues_WithRealValues()
    {
        var rep = PrivacyShield.Scan(File.ReadAllBytes(BuildDirty()));
        string V(string id) => rep.Findings.Single(f => f.Id == id).Value;
        Assert.Equal(Name1, V("author")); Assert.Equal(Name2, V("lastAuthor")); Assert.Equal(Comp, V("company")); Assert.Equal(Mgr, V("manager")); Assert.Equal(Prop, V("custom:Cliente"));
        Assert.Contains(rep.Findings, f => f.Category == "Autori di commenti e revisioni" && f.Value == Name1);
        Assert.Contains(rep.Findings, f => f.Category == "Percorsi locali" && f.Value.Contains(@"C:\Users\mrossi"));
        Assert.Contains(rep.Findings, f => f.Id == "thumbnail"); Assert.Contains(rep.Findings, f => f.Id == "comments");
        Assert.NotEmpty(rep.Warnings);   // le riserve ci sono sempre
    }

    [Fact]
    public void Clean_Anonymize_RemovesAllKnownData_OriginalUntouched_CopyValidAndSameContent()
    {
        var src = BuildDirty(); var orig = File.ReadAllBytes(src); var sha = Sha(orig);
        var rep = PrivacyShield.Scan(orig);
        var (clean, removed) = PrivacyShield.Clean(orig, rep.Findings.Select(f => f.Id).ToHashSet(), CommentMode.Anonymize);
        Assert.Equal(sha, Sha(File.ReadAllBytes(src)));   // l'originale è identico byte per byte
        Assert.Equal(sha, Sha(orig));                      // anche l'array in ingresso
        var all = AllText(clean);
        foreach (var secret in new[] { Name1, Name2, Comp, Mgr, Prop, "mrossi", @"C:\Users" }) Assert.DoesNotContain(secret, all);   // scansione part per part
        Assert.DoesNotContain("docProps/thumbnail", all);
        var after = PrivacyShield.Scan(clean);
        Assert.DoesNotContain(after.Findings, f => f.Id is "author" or "lastAuthor" or "company" or "manager" || f.Id.StartsWith("custom:") || f.Id.StartsWith("person:") || f.Id.StartsWith("path:") || f.Id == "thumbnail");
        Assert.NotEmpty(removed);
        Assert.Contains(after.Findings, f => f.Id == "comments");   // i commenti restano (anonimizzati)
        var copy = Path.Combine(Env.Work, "privacy-clean.pptx"); File.WriteAllBytes(copy, clean);
        Assert.Empty(RoundTripTests.NewErrors(src, copy));
        using var d1 = PptxDocument.Open(src); using var d2 = PptxDocument.Open(copy);
        Assert.Equal(d1.Slides.Count, d2.Slides.Count);
        string T(PptxDocument d) => string.Join("|", new Core.Rendering.SlideRenderer(d).Render(0).Items.OfType<Core.Rendering.TextItem>().Select(t => t.Text));
        Assert.Equal(T(d1), T(d2));
    }

    [Fact]
    public void Clean_RemoveComments_DeletesThem_AndFileStaysValid()
    {
        var src = BuildDirty(); var orig = File.ReadAllBytes(src);
        var (clean, _) = PrivacyShield.Clean(orig, PrivacyShield.Scan(orig).Findings.Select(f => f.Id).ToHashSet(), CommentMode.Remove);
        var copy = Path.Combine(Env.Work, "privacy-nocomments.pptx"); File.WriteAllBytes(copy, clean);
        using var z = ZipFile.OpenRead(copy);
        Assert.DoesNotContain(z.Entries, e => e.FullName.StartsWith("ppt/comments/") || e.FullName == "ppt/authors.xml");
        Assert.DoesNotContain(Name1, AllText(clean));
        Assert.Empty(RoundTripTests.NewErrors(src, copy));
    }

    [Fact]
    public void AlreadyCleanFile_FindsNothing_AndPartialSelectionLeavesTheRest()
    {
        var orig = File.ReadAllBytes(BuildDirty());
        var (clean, _) = PrivacyShield.Clean(orig, PrivacyShield.Scan(orig).Findings.Select(f => f.Id).ToHashSet(), CommentMode.Remove);
        Assert.Empty(PrivacyShield.Scan(clean).Findings);   // «niente trovato»
        var (partial, _) = PrivacyShield.Clean(orig, new HashSet<string> { "company" }, CommentMode.Anonymize);   // solo ciò che è spuntato
        var all = AllText(partial); Assert.DoesNotContain(Comp, all); Assert.Contains(Name1, all); Assert.Contains(Prop, all);
    }

    [Fact]
    public void WriteCopy_NeverOverwritesTheOriginal_AndIsAtomic()
    {
        var src = BuildDirty(); var bytes = File.ReadAllBytes(src);
        Assert.Throws<InvalidOperationException>(() => PrivacyShield.WriteCopy(bytes, src, src));
        Assert.Equal(bytes, File.ReadAllBytes(src));
        var dst = Path.Combine(Env.Work, "privacy-out (pulito).pptx"); PrivacyShield.WriteCopy(bytes, dst, src);
        Assert.Equal("privacy-dirty (pulito).pptx", PrivacyShield.ProposeName(src));
        Assert.Empty(Directory.GetFiles(Env.Work, "privacy-out*.tmp*"));
    }
}

[Collection("Desktop"), Trait("Category", "OfficeCom")]
public class PrivacyOfficeTests
{
    [Fact]
    public void CleanCopy_OpensInRealOffice_WithoutRepair_SameContent()
    {
        var src = PrivacyTests.BuildDirty(); var orig = File.ReadAllBytes(src);
        var (clean, _) = PrivacyShield.Clean(orig, PrivacyShield.Scan(orig).Findings.Select(f => f.Id).ToHashSet(), CommentMode.Anonymize);
        var copy = Path.Combine(Env.Work, "privacy-office.pptx"); File.WriteAllBytes(copy, clean);
        var r = ComRunner.Inspect(new[] { src, copy });
        Assert.True(ComRunner.Ok(r[src]), ComRunner.Error(r[src])); Assert.True(ComRunner.Ok(r[copy]), "Office apre la copia: " + ComRunner.Error(r[copy]));
        Assert.Equal(ComRunner.Signature(r[src]), ComRunner.Signature(r[copy]));
    }
}
