using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Core;

/// <summary>Part collegata a una forma copiata (immagine, grafico, media...) con i suoi figli: i byte viaggiano nella clip, così si incolla anche in un altro documento.</summary>
public sealed class ClipPart
{
    /// <summary>part | data (media) | hyper | ext</summary>
    public string Kind { get; set; } = "part";
    public string RelType { get; set; } = "";
    public string ContentType { get; set; } = "";
    public string Ext { get; set; } = "";
    public string? Uri { get; set; }
    public bool External { get; set; }
    public byte[]? Data { get; set; }
    public Dictionary<string, ClipPart> Children { get; set; } = new();
}

/// <summary>Forme copiate: XML di ciascuna + part collegate (chiave = id di relazione nella slide di origine) + testo semplice.</summary>
public sealed class ShapeClip
{
    public List<string> Shapes { get; set; } = new();
    public Dictionary<string, ClipPart> Parts { get; set; } = new();
    public string Text { get; set; } = "";

    public byte[] ToBytes() => JsonSerializer.SerializeToUtf8Bytes(this);

    /// <summary>Null se i byte non sono una clip valida (la clipboard è un input non fidato).</summary>
    public static ShapeClip? FromBytes(byte[]? b)
    {
        if (b == null || b.Length == 0) return null;
        try { var c = JsonSerializer.Deserialize<ShapeClip>(b); return c is { Shapes.Count: > 0 } ? c : null; }
        catch (JsonException) { return null; }
    }
}

public sealed partial class PptxDocument
{
    const long MaxClipBytes = 256L << 20;
    static readonly HashSet<string> ShapeNames = new() { "sp", "pic", "graphicFrame", "grpSp", "cxnSp" };
    static readonly string[] SafeSchemes = { "http", "https", "mailto" };

    static XElement SafeParse(string xml)
    {
        using var r = XmlReader.Create(new StringReader(xml), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        return XElement.Load(r);
    }

    OpenXmlElement? TopLevel(SlidePart sp, int id) =>
        sp.Slide!.CommonSlideData!.ShapeTree!.Elements().FirstOrDefault(e => ShapeNames.Contains(e.LocalName) && e.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value == (uint)id);

    /// <summary>Id di tutte le forme di primo livello della slide, in ordine di disegno.</summary>
    public List<int> ShapeIds(int slide) =>
        Slide(slide).Slide!.CommonSlideData!.ShapeTree!.Elements().Where(e => ShapeNames.Contains(e.LocalName)).Select(e => (int)(e.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value ?? 0)).Where(i => i != 0).ToList();

    /// <summary>Etichetta di una forma per elenchi (riquadro di selezione): nome e inizio del testo.</summary>
    public string ShapeLabel(int slide, int id)
    {
        var el = TopLevel(Slide(slide), id); if (el == null) return $"Forma {id}";
        var name = el.Descendants<P.NonVisualDrawingProperties>().First().Name?.Value ?? $"Forma {id}";
        var t = string.Concat(el.Descendants<A.Text>().Select(x => x.Text + " ")).Trim();
        return t.Length == 0 ? name : $"{name}: {(t.Length > 40 ? t[..40] + "…" : t)}";
    }

    static ClipPart Capture(OpenXmlPart p)
    {
        var cp = new ClipPart { Kind = "part", RelType = p.RelationshipType, ContentType = p.ContentType, Ext = Path.GetExtension(p.Uri.OriginalString) };
        using (var s = p.GetStream(FileMode.Open, FileAccess.Read)) { var ms = new MemoryStream(); s.CopyTo(ms); cp.Data = ms.ToArray(); }
        foreach (var c in p.Parts) cp.Children[c.RelationshipId] = Capture(c.OpenXmlPart);
        foreach (var h in p.ExternalRelationships) cp.Children[h.Id] = new ClipPart { Kind = "ext", RelType = h.RelationshipType, Uri = h.Uri.OriginalString, External = true };
        return cp;
    }

    /// <summary>Copia le forme indicate (con part collegate). I segnaposto diventano forme libere con posizione esplicita (ponytail: perdono la formattazione ereditata dal layout → copiarla nel file).</summary>
    public ShapeClip CopyShapes(int slide, IEnumerable<int> ids)
    {
        var sp = Slide(slide); var clip = new ShapeClip(); var texts = new List<string>();
        var boxes = new Lazy<List<(int Id, RectD Box, bool HasText)>>(() => new SlideRenderer(this).Render(slide).Hits);
        long total = 0;
        foreach (var id in ids)
        {
            var el = TopLevel(sp, id) ?? throw new KeyNotFoundException($"Forma {id} non trovata.");
            var x = SafeParse(el.OuterXml);
            var ph = x.Element(Ns.P + "nvSpPr")?.Element(Ns.P + "nvPr")?.Element(Ns.P + "ph");
            if (ph != null)
            {
                ph.Remove();
                var spPr = x.Element(Ns.P + "spPr")!;
                if (spPr.Element(Ns.A + "xfrm") == null)
                {
                    var b = boxes.Value.First(h => h.Id == id).Box;
                    spPr.AddFirst(new XElement(Ns.A + "xfrm", new XElement(Ns.A + "off", new XAttribute("x", Emu(b.X)), new XAttribute("y", Emu(b.Y))), new XElement(Ns.A + "ext", new XAttribute("cx", Emu(b.W)), new XAttribute("cy", Emu(b.H)))));
                }
            }
            clip.Shapes.Add(x.ToString(SaveOptions.DisableFormatting));
            texts.Add(string.Join("\n", el.Descendants<A.Paragraph>().Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text)))));
            foreach (var rid in x.DescendantsAndSelf().SelectMany(e => e.Attributes()).Where(a => a.Name.Namespace == Ns.R).Select(a => a.Value).Distinct())
            {
                if (clip.Parts.ContainsKey(rid)) continue;
                ClipPart? cp = null;
                if (sp.TryGetPartById(rid, out var part)) cp = Capture(part);
                else if (sp.HyperlinkRelationships.FirstOrDefault(h => h.Id == rid) is { } h) cp = new ClipPart { Kind = "hyper", RelType = "hyperlink", Uri = h.Uri.OriginalString, External = h.IsExternal };
                else if (sp.ExternalRelationships.FirstOrDefault(e => e.Id == rid) is { } er) cp = new ClipPart { Kind = "ext", RelType = er.RelationshipType, Uri = er.Uri.OriginalString, External = true };
                else if (sp.DataPartReferenceRelationships.FirstOrDefault(d => d.Id == rid) is { } dr)
                {
                    using var s = dr.DataPart.GetStream(FileMode.Open, FileAccess.Read); var ms = new MemoryStream(); s.CopyTo(ms);
                    cp = new ClipPart { Kind = "data", Uri = dr.DataPart.Uri.OriginalString, RelType = dr.RelationshipType, ContentType = dr.DataPart.ContentType, Ext = Path.GetExtension(dr.DataPart.Uri.OriginalString), Data = ms.ToArray() };
                }
                if (cp != null) { clip.Parts[rid] = cp; total += Size(cp); }
            }
            if (total > MaxClipBytes) throw new PptxFormatException("Le forme da copiare sono troppo grandi.");
        }
        clip.Text = string.Join("\n", texts.Where(t => t.Length > 0));
        return clip;
    }

    static long Size(ClipPart p) => (p.Data?.LongLength ?? 0) + p.Children.Values.Sum(Size);

    static void Rebuild(OpenXmlPartContainer parent, ClipPart c, string rid)
    {
        if (c.Kind == "ext")
        {
            if (c.Uri != null && Uri.TryCreate(c.Uri, UriKind.Absolute, out var u) && SafeSchemes.Contains(u.Scheme)) parent.AddExternalRelationship(c.RelType, u, rid);
            return;
        }
        if (c.Kind != "part") return;
        // part tipizzate (cartelle e nomi come li scrive PowerPoint); per le altre part una generica
        string t = c.RelType;
        OpenXmlPart np = t switch
        {
            _ when t.EndsWith("/image") => parent.AddNewPart<ImagePart>(c.ContentType, rid),
            _ when t.EndsWith("/chart") => parent.AddNewPart<ChartPart>(rid),
            _ when t.EndsWith("/chartStyle") => parent.AddNewPart<ChartStylePart>(rid),
            _ when t.EndsWith("/chartColorStyle") => parent.AddNewPart<ChartColorStylePart>(rid),
            _ when t.EndsWith("/diagramData") => parent.AddNewPart<DiagramDataPart>(rid),
            _ when t.EndsWith("/diagramLayout") => parent.AddNewPart<DiagramLayoutDefinitionPart>(rid),
            _ when t.EndsWith("/diagramQuickStyle") => parent.AddNewPart<DiagramStylePart>(rid),
            _ when t.EndsWith("/diagramColors") => parent.AddNewPart<DiagramColorsPart>(rid),
            _ => parent.AddExtendedPart(t, c.ContentType, c.Ext.Length > 0 ? c.Ext : ".xml", rid)
        };
        using (var s = new MemoryStream(c.Data ?? Array.Empty<byte>())) np.FeedData(s);
        foreach (var (crid, child) in c.Children) Rebuild(np, child, crid);
    }

    void BuildPart(SlidePart sp, ClipPart part, string rid, Dictionary<string, MediaDataPart> media)
    {
        switch (part.Kind)
        {
            case "hyper":
                if (part.Uri != null && Uri.TryCreate(part.Uri, UriKind.RelativeOrAbsolute, out var u) && (!part.External || (u.IsAbsoluteUri && SafeSchemes.Contains(u.Scheme)))) sp.AddHyperlinkRelationship(u, part.External, rid);
                break;
            case "data":
                {
                    // video/audio e relazione "media" puntano alla stessa part, come in PowerPoint
                    if (part.Uri == null || !media.TryGetValue(part.Uri, out var dp))
                    {
                        dp = _doc.CreateMediaDataPart(part.ContentType, part.Ext);
                        using (var s = new MemoryStream(part.Data ?? Array.Empty<byte>())) dp.FeedData(s);
                        if (part.Uri != null) media[part.Uri] = dp;
                    }
                    if (part.RelType.EndsWith("/video")) sp.AddVideoReferenceRelationship(dp, rid);
                    else if (part.RelType.EndsWith("/audio")) sp.AddAudioReferenceRelationship(dp, rid);
                    else sp.AddMediaReferenceRelationship(dp, rid);
                    break;
                }
            default: Rebuild(sp, part, rid); break;
        }
    }

    static (long x, long y)? FirstOff(XElement x) =>
        x.DescendantsAndSelf(Ns.A + "off").FirstOrDefault() is { } o && long.TryParse((string?)o.Attribute("x"), out var a) && long.TryParse((string?)o.Attribute("y"), out var b) ? (a, b) : null;

    /// <summary>Incolla le forme nella slide. Ritorna gli id nuovi (la prima per ogni forma). Se una forma sta già nella stessa posizione si sposta di 12 pt (come PowerPoint).</summary>
    public List<int> PasteShapes(int slide, ShapeClip clip)
    {
        var sp = Slide(slide); var tree = Tree(sp); var result = new List<int>();
        int next = NextShapeId(sp); int counter = 0;
        var xs = clip.Shapes.Select(SafeParse).ToList();
        foreach (var x in xs) if (!ShapeNames.Contains(x.Name.LocalName) || x.Name.Namespace != Ns.P) throw new PptxFormatException("Dati incollati non validi.");
        var taken = tree.Elements().Where(e => ShapeNames.Contains(e.LocalName)).Select(e => FirstOff(SafeParse(e.OuterXml))).Where(o => o != null).Select(o => o!.Value).ToHashSet();
        long shift = 0;
        if (xs.Count > 0 && FirstOff(xs[0]) is { } o0) while (taken.Contains((o0.x + shift, o0.y + shift)) && shift < 12 * 12700 * 20) shift += 12 * 12700;
        var map = new Dictionary<string, string>();
        string NewRid(string old)
        {
            if (map.TryGetValue(old, out var n)) return n;
            do n = "rIdPp" + (++counter); while (sp.TryGetPartById(n, out _) || sp.HyperlinkRelationships.Any(h => h.Id == n) || sp.ExternalRelationships.Any(h => h.Id == n) || sp.DataPartReferenceRelationships.Any(h => h.Id == n));
            return map[old] = n;
        }
        var built = new HashSet<string>(); var mediaParts = new Dictionary<string, MediaDataPart>();
        foreach (var x in xs)
        {
            string firstId = "";
            var idMap = new Dictionary<string, string>();
            foreach (var c in x.DescendantsAndSelf().Where(e => e.Name.LocalName == "cNvPr" && e.Name.Namespace == Ns.P)) { var old = (string?)c.Attribute("id") ?? "0"; var nid = (next++).ToString(); if (firstId.Length == 0) firstId = nid; idMap[old] = nid; c.SetAttributeValue("id", nid); }
            foreach (var c in x.Descendants().Where(e => e.Name.LocalName is "stCxn" or "endCxn").ToList()) { if (idMap.TryGetValue((string?)c.Attribute("id") ?? "", out var n)) c.SetAttributeValue("id", n); else c.Remove(); }
            if (shift != 0 && x.DescendantsAndSelf(Ns.A + "off").FirstOrDefault() is { } off) { off.SetAttributeValue("x", (long)off.Attribute("x")! + shift); off.SetAttributeValue("y", (long)off.Attribute("y")! + shift); }
            foreach (var a in x.DescendantsAndSelf().SelectMany(e => e.Attributes()).Where(a => a.Name.Namespace == Ns.R && a.Value.Length > 0).ToList())
            {
                if (clip.Parts.TryGetValue(a.Value, out var part))
                {
                    var nr = NewRid(a.Value);
                    if (built.Add(a.Value)) BuildPart(sp, part, nr, mediaParts);
                    a.Value = nr;
                }
                else if (a.Parent!.Name.LocalName is "hlinkClick" or "hlinkHover") a.Parent.Remove();
                else throw new PptxFormatException("Dati incollati non validi: relazione mancante.");
            }
            // collegamenti la cui relazione è stata scartata (schema non sicuro)
            foreach (var h in x.Descendants().Where(e => e.Name.LocalName is "hlinkClick" or "hlinkHover").ToList())
                if ((string?)h.Attribute(Ns.R + "id") is { Length: > 0 } rr && !sp.HyperlinkRelationships.Any(y => y.Id == rr) && !sp.ExternalRelationships.Any(y => y.Id == rr)) h.Remove();
            var xml = x.ToString(SaveOptions.DisableFormatting);
            OpenXmlElement el = x.Name.LocalName switch
            {
                "sp" => new P.Shape(xml), "pic" => new P.Picture(xml), "graphicFrame" => new P.GraphicFrame(xml), "grpSp" => new P.GroupShape(xml),
                _ => new P.ConnectionShape(xml)
            };
            tree.AppendChild(el);
            result.Add(int.Parse(firstId));
        }
        Dirty = true;
        return result;
    }

    /// <summary>Identificatore stabile della slide (resta uguale dopo spostamenti e duplicazioni di altre slide).</summary>
    public uint SlideUid(int i) => SlideIds!.Elements<P.SlideId>().ElementAt(i).Id!.Value;
    public int IndexOfUid(uint uid) => SlideIds!.Elements<P.SlideId>().ToList().FindIndex(s => s.Id?.Value == uid);
}
