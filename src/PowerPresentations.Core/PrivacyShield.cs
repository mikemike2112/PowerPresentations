using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace PowerPresentations.Core;

public sealed record PrivacyFinding(string Id, string Category, string Label, string Value, bool Cleanable = true);

public sealed class PrivacyReport
{
    public List<PrivacyFinding> Findings { get; } = new();
    /// <summary>Ciò che lo scudo NON sa pulire: va sempre mostrato all'utente.</summary>
    public List<string> Warnings { get; } = new();
}

public enum CommentMode { Anonymize, Remove }

/// <summary>
/// Scudo privacy: analizza un .pptx (in memoria) e produce una COPIA ripulita. Non tocca mai il file aperto.
/// Non dichiara mai "documento pulito": elenca sempre ciò che non sa trattare.
/// </summary>
public static class PrivacyShield
{
    static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    static readonly Regex AnonName = new(@"^Autore \d+$|^anonimo-\d+$", RegexOptions.Compiled);
    static readonly Regex LocalPath = new(@"(?<![\w])(?:[A-Za-z]:\\|\\\\[\w.\-]+\\)[^<>""|?*\r\n]+", RegexOptions.Compiled);

    static string Read(ZipArchiveEntry e) { using var r = new StreamReader(e.Open(), Encoding.UTF8); return r.ReadToEnd(); }

    static XDocument? TryXml(ZipArchive z, string name)
    {
        var e = z.GetEntry(name); if (e == null) return null;
        try { using var s = e.Open(); return XDocument.Load(s); } catch (System.Xml.XmlException) { return null; }
    }

    static bool IsAuthorPart(string n) => Regex.IsMatch(n, @"^ppt/(commentAuthors|authors|people)\.xml$", RegexOptions.IgnoreCase);
    static readonly string[] PersonAttrs = { "name", "displayName", "userId", "initials" };

    // ---------- analisi ----------
    public static PrivacyReport Scan(byte[] pptx)
    {
        var rep = new PrivacyReport();
        using var ms = new MemoryStream(pptx, false); using var z = new ZipArchive(ms, ZipArchiveMode.Read);
        var core = TryXml(z, "docProps/core.xml");
        void Add(string id, string cat, string label, string? v) { if (!string.IsNullOrWhiteSpace(v) && !AnonName.IsMatch(v.Trim())) rep.Findings.Add(new PrivacyFinding(id, cat, label, v.Trim())); }
        Add("author", "Metadati", "Autore", core?.Descendants(Dc + "creator").FirstOrDefault()?.Value);
        Add("lastAuthor", "Metadati", "Ultimo autore", core?.Descendants(Cp + "lastModifiedBy").FirstOrDefault()?.Value);
        var app = TryXml(z, "docProps/app.xml");
        Add("company", "Metadati", "Società", app?.Descendants().FirstOrDefault(e => e.Name.LocalName == "Company")?.Value);
        Add("manager", "Metadati", "Manager", app?.Descendants().FirstOrDefault(e => e.Name.LocalName == "Manager")?.Value);
        var custom = TryXml(z, "docProps/custom.xml");
        foreach (var p in custom?.Descendants().Where(e => e.Name.LocalName == "property") ?? Enumerable.Empty<XElement>())
            rep.Findings.Add(new PrivacyFinding("custom:" + (string?)p.Attribute("name"), "Proprietà personalizzate", (string?)p.Attribute("name") ?? "(senza nome)", p.Value.Trim()));
        // autori di commenti e revisioni
        foreach (var e in z.Entries.Where(e => IsAuthorPart(e.FullName)))
        {
            var x = TryXml(z, e.FullName); if (x == null) continue;
            foreach (var el in x.Descendants())
                foreach (var a in el.Attributes().Where(a => PersonAttrs.Contains(a.Name.LocalName)))
                    if (!string.IsNullOrWhiteSpace(a.Value) && !AnonName.IsMatch(a.Value) && a.Name.LocalName is "name" or "displayName")
                        rep.Findings.Add(new PrivacyFinding($"person:{e.FullName}:{a.Value}", "Autori di commenti e revisioni", "Autore (" + e.FullName.Split('/').Last() + ")", a.Value));
        }
        int comments = z.Entries.Count(e => Regex.IsMatch(e.FullName, @"^ppt/comments/.+\.xml$", RegexOptions.IgnoreCase));
        if (comments > 0) rep.Findings.Add(new PrivacyFinding("comments", "Commenti", "Commenti nelle slide", $"{comments} elenchi di commenti (testo e data restano se si anonimizza)"));
        // percorsi locali e nomi utente nei metadati
        foreach (var e in z.Entries.Where(e => e.FullName.StartsWith("docProps/") && e.FullName.EndsWith(".xml")))
            foreach (Match m in LocalPath.Matches(Read(e))) rep.Findings.Add(new PrivacyFinding($"path:{e.FullName}:{m.Value}", "Percorsi locali", "Percorso in " + e.FullName.Split('/').Last(), m.Value));
        if (z.Entries.Any(e => e.FullName.StartsWith("docProps/thumbnail", StringComparison.OrdinalIgnoreCase)))
            rep.Findings.Add(new PrivacyFinding("thumbnail", "Anteprima", "Miniatura del file", "docProps/thumbnail (mostra la prima slide)"));

        // avvisi: ciò che non si sa pulire
        int exif = 0;
        foreach (var e in z.Entries.Where(e => e.FullName.StartsWith("ppt/media/") && Regex.IsMatch(e.FullName, @"\.jpe?g$", RegexOptions.IgnoreCase)))
        {
            var buf = new byte[Math.Min(65536, (int)Math.Min(e.Length, 65536))]; using var s = e.Open(); int n = s.Read(buf, 0, buf.Length);
            if (Encoding.Latin1.GetString(buf, 0, n).Contains("Exif\0\0")) exif++;
        }
        if (exif > 0) rep.Warnings.Add($"{exif} immagini JPEG contengono metadati EXIF (posizione, fotocamera, data): NON vengono ripuliti.");
        int emb = z.Entries.Count(e => e.FullName.StartsWith("ppt/embeddings/")); if (emb > 0) rep.Warnings.Add($"{emb} oggetti incorporati (fogli, documenti…): il loro contenuto e i loro metadati NON vengono analizzati.");
        if (z.Entries.Any(e => e.FullName.Contains("vbaProject", StringComparison.OrdinalIgnoreCase))) rep.Warnings.Add("Il file contiene macro: NON vengono analizzate né rimosse.");
        int hidden = 0, notes = 0;
        foreach (var e in z.Entries.Where(e => Regex.IsMatch(e.FullName, @"^ppt/slides/slide\d+\.xml$"))) { var t = Read(e); if (Regex.IsMatch(t, @"<p:sld [^>]*\bshow=""(0|false)""")) hidden++; }
        foreach (var e in z.Entries.Where(e => Regex.IsMatch(e.FullName, @"^ppt/notesSlides/notesSlide\d+\.xml$"))) if (Regex.IsMatch(Read(e), @"<a:t>[^<]+</a:t>")) notes++;
        if (hidden > 0) rep.Warnings.Add($"{hidden} slide nascoste: il loro testo NON viene rimosso.");
        if (notes > 0) rep.Warnings.Add($"{notes} slide hanno note del relatore: il testo delle note NON viene analizzato.");
        int ext = 0;
        foreach (var e in z.Entries.Where(e => e.FullName.EndsWith(".rels")))
            ext += Regex.Matches(Read(e), @"TargetMode=""External""[^>]*Target=""(file:|[A-Za-z]:\\|\\\\)|Target=""(file:|[A-Za-z]:\\|\\\\)[^""]*""[^>]*TargetMode=""External""").Count;
        if (ext > 0) rep.Warnings.Add($"{ext} collegamenti a file locali o di rete: NON vengono modificati (sono funzionali).");
        rep.Warnings.Add("Il testo delle slide, i nomi di file delle immagini e altri contenuti scritti dall'utente non vengono analizzati.");
        return rep;
    }

    // ---------- pulizia della copia ----------
    static void Edit(ZipArchive z, string name, Func<XDocument, bool> change)
    {
        var e = z.GetEntry(name); if (e == null) return;
        XDocument x; using (var s = e.Open()) x = XDocument.Load(s);
        if (!change(x)) return;
        e.Delete(); var ne = z.CreateEntry(name);
        using var w = new StreamWriter(ne.Open(), new UTF8Encoding(false)); w.Write(x.Declaration != null ? x.Declaration + "\n" : "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"); w.Write(x.Root!.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>Ritorna la copia ripulita e l'elenco di ciò che è stato tolto. L'array in ingresso non viene modificato.</summary>
    public static (byte[] Data, List<string> Removed) Clean(byte[] pptx, ISet<string> ids, CommentMode comments = CommentMode.Anonymize)
    {
        var removed = new List<string>();
        var data = (byte[])pptx.Clone();
        // 1) cancellazioni strutturali con l'SDK (relazioni e content type coerenti)
        bool removeComments = ids.Contains("comments") && comments == CommentMode.Remove, removeThumb = ids.Contains("thumbnail");
        if (removeComments || removeThumb)
        {
            using var ms = new MemoryStream(); ms.Write(data); ms.Position = 0;
            using (var doc = PresentationDocument.Open(ms, true, new OpenSettings { AutoSave = false }))
            {
                if (removeThumb && doc.ThumbnailPart != null) { doc.DeletePart(doc.ThumbnailPart); removed.Add("Miniatura del file"); }
                if (removeComments)
                {
                    var pp = doc.PresentationPart!;
                    foreach (var sp in pp.SlideParts)
                    {
                        foreach (var c in sp.GetPartsOfType<SlideCommentsPart>().ToList()) sp.DeletePart(c);
                        foreach (var c in sp.GetPartsOfType<PowerPointCommentPart>().ToList()) sp.DeletePart(c);
                    }
                    foreach (var a in pp.GetPartsOfType<CommentAuthorsPart>().ToList()) pp.DeletePart(a);
                    foreach (var a in pp.GetPartsOfType<PowerPointAuthorsPart>().ToList()) pp.DeletePart(a);
                    removed.Add("Commenti e relativi autori (rimossi)");
                }
                doc.Save();
            }
            data = ms.ToArray();
        }
        // 2) modifiche testuali mirate: le altre part restano byte per byte come sono
        using (var ms = new MemoryStream())
        {
            ms.Write(data); ms.Position = 0;
            using (var z = new ZipArchive(ms, ZipArchiveMode.Update, true))
            {
                Edit(z, "docProps/core.xml", x =>
                {
                    bool ch = false;
                    foreach (var (id, el, label) in new[] { ("author", Dc + "creator", "Autore"), ("lastAuthor", Cp + "lastModifiedBy", "Ultimo autore") })
                        if (ids.Contains(id)) foreach (var e in x.Descendants(el).Where(e => e.Value.Length > 0)) { e.Value = ""; ch = true; removed.Add(label); }
                    return ch;
                });
                Edit(z, "docProps/app.xml", x =>
                {
                    bool ch = false;
                    foreach (var (id, nm, label) in new[] { ("company", "Company", "Società"), ("manager", "Manager", "Manager") })
                        if (ids.Contains(id)) foreach (var e in x.Descendants().Where(e => e.Name.LocalName == nm && e.Value.Length > 0)) { e.Value = ""; ch = true; removed.Add(label); }
                    return ch;
                });
                Edit(z, "docProps/custom.xml", x =>
                {
                    var del = x.Descendants().Where(e => e.Name.LocalName == "property" && ids.Contains("custom:" + (string?)e.Attribute("name"))).ToList();
                    foreach (var e in del) { removed.Add("Proprietà personalizzata «" + (string?)e.Attribute("name") + "»"); e.Remove(); }
                    return del.Count > 0;
                });
                // percorsi locali nei metadati: si tiene solo il nome del file
                foreach (var e in z.Entries.Where(e => e.FullName.StartsWith("docProps/") && e.FullName.EndsWith(".xml")).ToList())
                    Edit(z, e.FullName, x =>
                    {
                        bool ch = false;
                        foreach (var t in x.DescendantNodes().OfType<XText>().ToList())
                        {
                            var nv = LocalPath.Replace(t.Value, m => ids.Contains($"path:{e.FullName}:{m.Value}") ? Path.GetFileName(m.Value.TrimEnd('\\')) : m.Value);
                            if (nv != t.Value) { t.Value = nv; ch = true; removed.Add("Percorso locale in " + e.FullName.Split('/').Last()); }
                        }
                        return ch;
                    });
                // autori (anonimizzazione coerente: stesso nome → stesso «Autore N»)
                var map = new Dictionary<string, string>(StringComparer.Ordinal);
                string Anon(string v) { if (!map.TryGetValue(v, out var a)) map[v] = a = "Autore " + (map.Count + 1); return a; }
                foreach (var name in z.Entries.Where(e => IsAuthorPart(e.FullName)).Select(e => e.FullName).ToList())
                    Edit(z, name, x =>
                    {
                        bool ch = false;
                        foreach (var el in x.Descendants())
                        {
                            var nm = el.Attributes().FirstOrDefault(a => a.Name.LocalName is "name" or "displayName");
                            if (nm == null || !ids.Contains($"person:{name}:{nm.Value}") || AnonName.IsMatch(nm.Value)) continue;
                            string a1 = Anon(nm.Value), n = a1["Autore ".Length..];
                            foreach (var a in el.Attributes().ToList())
                                switch (a.Name.LocalName) { case "name": case "displayName": a.Value = a1; break; case "initials": a.Value = "A" + n; break; case "userId": a.Value = "anonimo-" + n; break; case "providerId": a.Value = "None"; break; }
                            ch = true;
                        }
                        if (ch) removed.Add($"Autori anonimizzati in {name.Split('/').Last()} ({map.Count})");
                        return ch;
                    });
            }
            data = ms.ToArray();
        }
        return (data, removed.Distinct().ToList());
    }

    /// <summary>Scrive la copia in modo atomico (file temporaneo + sostituzione). Mai sull'originale: se il percorso coincide, rifiuta.</summary>
    public static void WriteCopy(byte[] data, string path, string? originalPath)
    {
        if (originalPath != null && string.Equals(Path.GetFullPath(originalPath), Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("La copia ripulita non può sovrascrivere il file originale: scegli un altro nome.");
        var tmp = path + ".tmp" + Environment.ProcessId;
        try { File.WriteAllBytes(tmp, data); File.Move(tmp, path, true); }
        catch { try { if (File.Exists(tmp)) File.Delete(tmp); } catch (IOException) { } throw; }
    }

    public static string ProposeName(string? original) =>
        original == null ? "Presentazione (pulito).pptx" : Path.GetFileNameWithoutExtension(original) + " (pulito)" + Path.GetExtension(original);
}
