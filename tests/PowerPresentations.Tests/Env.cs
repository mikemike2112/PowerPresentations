using System.IO.Packaging;
using System.Text;
using System.Xml.Linq;

namespace PowerPresentations.Tests;

public static class Env
{
    public static readonly string Repo = FindRepo();
    public static string Corpus => Path.Combine(Repo, "tests", "corpus", "office-made");
    public static string Tools => Path.Combine(Repo, "tools");
    public static string Work { get; } = Fresh(Path.Combine(AppContext.BaseDirectory, "work"));

    static string Fresh(string d) { if (Directory.Exists(d)) Directory.Delete(d, true); return Directory.CreateDirectory(d).FullName; }

    static string FindRepo()
    {
        var d = new DirectoryInfo(AppContext.BaseDirectory);
        while (d != null && !File.Exists(Path.Combine(d.FullName, "PowerPresentations.sln"))) d = d.Parent;
        return d?.FullName ?? throw new InvalidOperationException("Radice della repo non trovata.");
    }

    public static IEnumerable<object[]> CorpusFiles() =>
        Directory.GetFiles(Corpus, "*.pptx").OrderBy(f => f).Select(f => new object[] { Path.GetFileName(f) });

    public static string CorpusPath(string name) => Path.Combine(Corpus, name);

    /// <summary>Copia di lavoro del file (il corpus non si modifica mai).</summary>
    public static string Copy(string name, string? suffix = null)
    {
        var dst = Path.Combine(Work, Path.GetFileNameWithoutExtension(name) + suffix + ".pptx");
        File.Copy(CorpusPath(name), dst, true); return dst;
    }
}

/// <summary>Confronto di package: stesso elenco di part e relazioni, XML canonicalizzato, binari identici.</summary>
public static class PackageCompare
{
    public static string Canon(XElement e) { var sb = new StringBuilder(); Write(e, sb); return sb.ToString(); }

    static void Write(XElement e, StringBuilder sb)
    {
        sb.Append('<').Append(e.Name.ToString());
        foreach (var a in e.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal))
            sb.Append(' ').Append(a.Name.ToString()).Append("=\"").Append(a.Value).Append('"');
        sb.Append('>');
        foreach (var n in e.Nodes())
        {
            if (n is XElement c) Write(c, sb);
            else if (n is XText t && t.Value.Trim().Length > 0) sb.Append(t.Value);
        }
        sb.Append("</").Append(e.Name.ToString()).Append('>');
    }

    static byte[] Bytes(PackagePart p) { using var s = p.GetStream(); using var m = new MemoryStream(); s.CopyTo(m); return m.ToArray(); }

    public static List<string> Diff(string a, string b)
    {
        var diffs = new List<string>();
        using var pa = Package.Open(a, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var pb = Package.Open(b, FileMode.Open, FileAccess.Read, FileShare.Read);
        var partsA = pa.GetParts().Where(p => !p.ContentType.EndsWith("relationships+xml")).ToDictionary(p => p.Uri.ToString());
        var partsB = pb.GetParts().Where(p => !p.ContentType.EndsWith("relationships+xml")).ToDictionary(p => p.Uri.ToString());
        foreach (var k in partsA.Keys.Except(partsB.Keys)) diffs.Add("part mancante: " + k);
        foreach (var k in partsB.Keys.Except(partsA.Keys)) diffs.Add("part in più: " + k);
        foreach (var k in partsA.Keys.Intersect(partsB.Keys))
        {
            var x = partsA[k]; var y = partsB[k];
            if (x.ContentType != y.ContentType) diffs.Add($"content type {k}: {x.ContentType} ≠ {y.ContentType}");
            if (Rels(x.Uri, x.GetRelationships()) != Rels(y.Uri, y.GetRelationships())) diffs.Add("relazioni diverse: " + k);
            var ba = Bytes(x); var bb = Bytes(y);
            if (x.ContentType.Contains("xml", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var ca = Canon(XElement.Parse(Encoding.UTF8.GetString(ba).TrimStart('﻿')));
                    var cb = Canon(XElement.Parse(Encoding.UTF8.GetString(bb).TrimStart('﻿')));
                    if (ca != cb) diffs.Add("XML diverso: " + k);
                }
                catch (System.Xml.XmlException) { if (!ba.SequenceEqual(bb)) diffs.Add("contenuto diverso: " + k); }
            }
            else if (!ba.SequenceEqual(bb)) diffs.Add("binario diverso: " + k);
        }
        if (Rels(new Uri("/", UriKind.Relative), pa.GetRelationships()) != Rels(new Uri("/", UriKind.Relative), pb.GetRelationships())) diffs.Add("relazioni del package diverse");
        return diffs;
    }

    static string Rels(Uri source, IEnumerable<PackageRelationship> r) =>
        string.Join("|", r.Select(x => $"{x.Id}>{(x.TargetMode == TargetMode.Internal ? PackUriHelper.ResolvePartUri(source, x.TargetUri).ToString() : x.TargetUri.ToString())}>{x.RelationshipType}>{x.TargetMode}").OrderBy(s => s, StringComparer.Ordinal));
}
