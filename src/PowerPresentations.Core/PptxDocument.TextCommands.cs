using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPresentations.Core;

public sealed record FindHit(int Slide, int ShapeId, string Text);

/// <summary>Formato copiato: proprietà della prima run, del primo paragrafo e riempimento/contorno della forma.</summary>
public sealed class FormatClip
{
    internal A.RunProperties? Run; internal A.ParagraphProperties? Para; internal List<OpenXmlElement> Fill = new(); internal bool HasFill;
}

/// <summary>Trova e sostituisci, collegamenti, simboli, copia formato.</summary>
public sealed partial class PptxDocument
{
    static Regex Pattern(string find, bool matchCase) => new(Regex.Escape(find), RegexOptions.CultureInvariant | (matchCase ? 0 : RegexOptions.IgnoreCase));

    /// <summary>Tutte le occorrenze nel testo di forme e celle (ponytail: una parola spezzata fra due run con formati diversi non si trova → unire le run).</summary>
    public List<FindHit> FindAll(string find, bool matchCase = false)
    {
        if (find.Length == 0) throw new ArgumentException("Testo da cercare vuoto.", nameof(find));
        var rx = Pattern(find, matchCase); var hits = new List<FindHit>();
        for (int i = 0; i < Slides.Count; i++)
            foreach (var el in Tree(Slide(i)).Elements().Where(e => ShapeNames.Contains(e.LocalName)))
            {
                int id = (int)(el.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value ?? 0);
                foreach (var t in el.Descendants<A.Text>()) if (rx.IsMatch(t.Text)) hits.Add(new FindHit(i, id, t.Text));
            }
        return hits;
    }

    /// <summary>Sostituisce in tutta la presentazione; il formato delle run resta. Ritorna il numero di sostituzioni.</summary>
    public int ReplaceAll(string find, string repl, bool matchCase = false)
    {
        if (find.Length == 0) throw new ArgumentException("Testo da cercare vuoto.", nameof(find));
        var rx = Pattern(find, matchCase); int n = 0;
        for (int i = 0; i < Slides.Count; i++)
            foreach (var t in Slide(i).Slide!.Descendants<A.Text>().ToList())
            {
                var nt = rx.Replace(t.Text, _ => { n++; return repl; });
                if (nt != t.Text) t.Text = nt;
            }
        if (n > 0) Dirty = true;
        return n;
    }

    // ---------- collegamento ipertestuale ----------
    /// <summary>Solo http, https, mailto: un indirizzo scritto dall'utente non deve poter lanciare file o programmi.</summary>
    public static Uri ValidateLink(string url)
    {
        url = url.Trim();
        var full = url.Contains("://") || url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ? url : "https://" + url;
        if (url.Length == 0 || url.Length > 2000 || !Uri.TryCreate(full, UriKind.Absolute, out var u) || !SafeSchemes.Contains(u.Scheme))
            throw new ArgumentException("Indirizzo non valido: sono ammessi solo http, https e mailto.");
        return u;
    }

    /// <summary>Collega tutto il testo della forma (o delle celle) all'indirizzo; indirizzo vuoto = toglie i collegamenti. ponytail: nessuna selezione parziale → intervallo di run.</summary>
    public void SetHyperlink(int slide, int id, string? url)
    {
        var sp = Slide(slide); var bodies = TextBodies(slide, id);
        Uri? u = string.IsNullOrWhiteSpace(url) ? null : ValidateLink(url);
        string? rid = null;
        if (u != null) rid = (sp.HyperlinkRelationships.FirstOrDefault(h => h.IsExternal && h.Uri == u) ?? sp.AddHyperlinkRelationship(u, true)).Id;
        foreach (var r in bodies.SelectMany(t => t.Descendants<A.Run>()))
        {
            var rp = r.RunProperties ??= new A.RunProperties { Language = "it-IT" };
            foreach (var h in rp.Elements<A.HyperlinkOnClick>().ToList()) h.Remove();
            if (rid == null) continue;
            var link = new A.HyperlinkOnClick { Id = rid };
            var anchor = rp.ChildElements.FirstOrDefault(c => c is A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList);
            if (anchor != null) anchor.InsertBeforeSelf(link); else rp.AppendChild(link);
        }
        // relazioni non più usate
        foreach (var h in sp.HyperlinkRelationships.ToList())
            if (!sp.Slide!.Descendants().Any(e => e.GetAttributes().Any(a => a.NamespaceUri.EndsWith("/relationships") && a.Value == h.Id))) sp.DeleteReferenceRelationship(h.Id);
        if (bodies.Count > 0) Dirty = true;
    }

    /// <summary>Indirizzo del collegamento sulla prima run che ne ha uno.</summary>
    public string? GetHyperlink(int slide, int id)
    {
        var sp = Slide(slide);
        var rid = TextBodies(slide, id).SelectMany(t => t.Descendants<A.HyperlinkOnClick>()).Select(h => h.Id?.Value).FirstOrDefault(i => i != null);
        return rid == null ? null : sp.HyperlinkRelationships.FirstOrDefault(h => h.Id == rid)?.Uri.OriginalString;
    }

    // ---------- simboli ----------
    /// <summary>Aggiunge testo in coda all'ultima run (stesso formato). Senza testo crea la run.</summary>
    public void AppendText(int slide, int id, string text)
    {
        var bodies = TextBodies(slide, id); if (bodies.Count == 0) throw new NotSupportedException("L'oggetto non contiene testo.");
        var last = bodies[0].Descendants<A.Run>().LastOrDefault();
        if (last?.Text != null) last.Text.Text += text;
        else
        {
            var p = bodies[0].Elements<A.Paragraph>().LastOrDefault() ?? bodies[0].AppendChild(new A.Paragraph());
            var run = new A.Run(new A.RunProperties { Language = "it-IT" }, new A.Text(text));
            var end = p.GetFirstChild<A.EndParagraphRunProperties>(); if (end != null) end.InsertBeforeSelf(run); else p.AppendChild(run);
        }
        Dirty = true;
    }

    // ---------- copia formato ----------
    static readonly Type[] FillTypes = { typeof(A.NoFill), typeof(A.SolidFill), typeof(A.GradientFill), typeof(A.BlipFill), typeof(A.PatternFill), typeof(A.GroupFill) };
    static bool IsFillOrLine(OpenXmlElement c) => FillTypes.Contains(c.GetType()) || c is A.Outline;

    public FormatClip CopyFormat(int slide, int id)
    {
        var fc = new FormatClip();
        var bodies = TextBodies(slide, id);
        var rp = bodies.SelectMany(t => t.Descendants<A.Run>()).FirstOrDefault()?.RunProperties;
        if (rp != null)
        {
            fc.Run = (A.RunProperties)rp.CloneNode(true);
            foreach (var h in fc.Run.ChildElements.Where(c => c is A.HyperlinkOnClick or A.HyperlinkOnMouseOver).ToList()) h.Remove();
        }
        var pp = bodies.SelectMany(t => t.Elements<A.Paragraph>()).FirstOrDefault()?.ParagraphProperties;
        if (pp != null) fc.Para = (A.ParagraphProperties)pp.CloneNode(true);
        if (FindShape(Slide(slide), id)?.ShapeProperties is { } sp)
        {
            fc.HasFill = true;
            foreach (var c in sp.ChildElements.Where(IsFillOrLine)) fc.Fill.Add(c.CloneNode(true));
        }
        return fc;
    }

    /// <summary>Applica il formato copiato a testo e forma di destinazione (testo, lingua e collegamenti restano).</summary>
    public void ApplyFormat(int slide, int id, FormatClip fc)
    {
        foreach (var tb in TextBodies(slide, id))
        {
            if (fc.Run != null)
                foreach (var r in tb.Descendants<A.Run>())
                {
                    var old = r.RunProperties; var lang = old?.Language?.Value;
                    var links = old?.Elements<A.HyperlinkOnClick>().Select(h => h.CloneNode(true)).ToList();
                    var nrp = (A.RunProperties)fc.Run.CloneNode(true);
                    if (lang != null) nrp.Language = lang;
                    r.RunProperties = nrp;
                    if (links != null)
                        foreach (var l in links)
                        {
                            var a = nrp.ChildElements.FirstOrDefault(c => c is A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList);
                            if (a != null) a.InsertBeforeSelf(l); else nrp.AppendChild(l);
                        }
                }
            if (fc.Para != null)
                foreach (var p in tb.Elements<A.Paragraph>()) p.ParagraphProperties = (A.ParagraphProperties)fc.Para.CloneNode(true);
        }
        if (fc.HasFill && FindShape(Slide(slide), id) is { } sh)
        {
            var spPr = sh.ShapeProperties ??= new P.ShapeProperties();
            foreach (var c in spPr.ChildElements.Where(IsFillOrLine).ToList()) c.Remove();
            // ordine di spPr: xfrm, geometria, riempimento, contorno, effetti...
            var anchor = spPr.ChildElements.LastOrDefault(c => c is A.Transform2D or A.PresetGeometry or A.CustomGeometry);
            foreach (var c in fc.Fill.Select(c => c.CloneNode(true)).OrderBy(c => c is A.Outline ? 1 : 0))
            {
                if (anchor != null) anchor.InsertAfterSelf(c); else spPr.InsertAt(c, 0);
                anchor = c;
            }
        }
        Dirty = true;
    }
}
