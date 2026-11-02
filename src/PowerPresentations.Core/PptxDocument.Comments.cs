using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Core;

/// <summary>Lettura, modifica ed eliminazione dei commenti: moderni (PowerPointCommentPart, con risposte) e classici (SlideCommentsPart, risposte via p15:parentCm).</summary>
public sealed partial class PptxDocument
{
    const string CommentRelExt = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}";

    static string CleanCommentText(string text)
    {
        text = text.Replace("\r\n", "\n").Trim();
        if (text.Length == 0) throw new ArgumentException("Il commento è vuoto.");
        if (text.Length > 10000) throw new ArgumentException("Il commento è troppo lungo (massimo 10000 caratteri).");
        return text;
    }

    static IEnumerable<XElement> CommentParas(string text) =>
        text.Split('\n').Select(l => new XElement(Ns.A + "p", l.Length == 0 ? null : new XElement(Ns.A + "r", new XElement(Ns.A + "rPr", new XAttribute("lang", "it-IT")), new XElement(Ns.A + "t", l))));

    static string BodyText(XElement? body) => body == null ? "" : string.Join("\n", body.Elements(Ns.A + "p").Select(p => string.Concat(p.Descendants(Ns.A + "t").Select(t => t.Value))));

    // classico: la risposta porta p15:parentCm (authorId, idx) del commento a cui risponde
    static (string Author, string Idx)? LegacyParent(XElement cm) =>
        cm.Descendants().FirstOrDefault(e => e.Name.LocalName == "parentCm") is { } pc ? ((string?)pc.Attribute("authorId") ?? "", (string?)pc.Attribute("idx") ?? "") : null;

    static string LegacyId(XElement cm) => (string?)cm.Attribute("authorId") + ":" + (string?)cm.Attribute("idx");

    /// <summary>Commenti della slide: per ogni commento le sue risposte subito dopo (IsReply).</summary>
    public IReadOnlyList<CommentInfo> GetComments(int slide)
    {
        var sp = Slide(slide); var list = new List<CommentInfo>();
        if (sp.GetPartsOfType<PowerPointCommentPart>().FirstOrDefault() is { } mp)
        {
            var authors = ModernAuthors();
            string Au(XElement e) => authors.GetValueOrDefault((string?)e.Attribute("authorId") ?? "", "?");
            foreach (var cm in Xml(mp).Elements(N188 + "cm"))
            {
                var replies = cm.Element(N188 + "replyLst")?.Elements(N188 + "reply").ToList() ?? new();
                list.Add(new CommentInfo(Au(cm), BodyText(cm.Element(N188 + "txBody")), (string?)cm.Attribute("created") ?? "", (string?)cm.Attribute("id") ?? "", false, replies.Count));
                foreach (var r in replies) list.Add(new CommentInfo(Au(r), BodyText(r.Element(N188 + "txBody")), (string?)r.Attribute("created") ?? "", (string?)r.Attribute("id") ?? "", true));
            }
        }
        if (sp.SlideCommentsPart is { } lp)
        {
            var authors = LegacyAuthors(); var all = Xml(lp).Elements(Ns.P + "cm").ToList();
            CommentInfo Info(XElement cm, bool reply, int n) => new(authors.GetValueOrDefault((string?)cm.Attribute("authorId") ?? "", "?"), cm.Element(Ns.P + "text")?.Value ?? "", (string?)cm.Attribute("dt") ?? "", LegacyId(cm), reply, n);
            foreach (var cm in all.Where(c => LegacyParent(c) == null))
            {
                var kids = all.Where(c => LegacyParent(c) is { } p && p.Author == (string?)cm.Attribute("authorId") && p.Idx == (string?)cm.Attribute("idx")).ToList();
                list.Add(Info(cm, false, kids.Count));
                foreach (var k in kids) list.Add(Info(k, true, 0));
            }
        }
        return list;
    }

    /// <summary>Cambia solo il testo: autore, data, posizione e gli altri commenti restano intatti.</summary>
    public void EditComment(int slide, string id, string text)
    {
        text = CleanCommentText(text);
        var sp = Slide(slide);
        if (sp.GetPartsOfType<PowerPointCommentPart>().FirstOrDefault() is { } mp)
        {
            var list = Xml(mp);
            if (list.Descendants().FirstOrDefault(e => (e.Name == N188 + "cm" || e.Name == N188 + "reply") && (string?)e.Attribute("id") == id)?.Element(N188 + "txBody") is { } body)
            {
                body.Elements(Ns.A + "p").Remove(); body.Add(CommentParas(text));
                WritePart(mp, list); Dirty = true; return;
            }
        }
        if (sp.SlideCommentsPart is { } lp)
        {
            var list = Xml(lp);
            if (list.Elements(Ns.P + "cm").FirstOrDefault(c => LegacyId(c) == id) is { } cm)
            {
                if (cm.Element(Ns.P + "text") is { } t) t.Value = text; else cm.Add(new XElement(Ns.P + "text", text));
                WritePart(lp, list); Dirty = true; return;
            }
        }
        throw new KeyNotFoundException("Commento non trovato.");
    }

    /// <summary>Toglie dalla slide l'estensione p188:commentRel (e l'elenco estensioni se resta vuoto).</summary>
    static void DropCommentRel(SlidePart sp)
    {
        if (sp.Slide!.GetFirstChild<P.SlideExtensionList>() is not { } el) return;
        foreach (var x in el.Elements<P.SlideExtension>().Where(x => x.Uri == CommentRelExt).ToList()) x.Remove();
        if (!el.HasChildren) el.Remove();
    }

    /// <summary>Elimina un commento con le sue risposte. Se era l'ultimo della slide, part e relazione spariscono (autori e lastIdx restano come in PowerPoint).</summary>
    public void DeleteComment(int slide, string id)
    {
        var sp = Slide(slide);
        if (sp.GetPartsOfType<PowerPointCommentPart>().FirstOrDefault() is { } mp)
        {
            var list = Xml(mp);
            var hit = list.Descendants().FirstOrDefault(e => (e.Name == N188 + "cm" || e.Name == N188 + "reply") && (string?)e.Attribute("id") == id);
            if (hit != null)
            {
                var lst = hit.Parent!; hit.Remove();
                if (lst.Name == N188 + "replyLst" && !lst.HasElements) lst.Remove();
                if (list.Elements(N188 + "cm").Any()) WritePart(mp, list);
                else
                {
                    DropCommentRel(sp);
                    sp.DeletePart(mp);
                }
                Dirty = true; return;
            }
        }
        if (sp.SlideCommentsPart is { } lp)
        {
            var list = Xml(lp); var all = list.Elements(Ns.P + "cm").ToList();
            if (all.FirstOrDefault(c => LegacyId(c) == id) is { } cm)
            {
                var (au, ix) = ((string?)cm.Attribute("authorId") ?? "", (string?)cm.Attribute("idx") ?? "");
                foreach (var k in all.Where(c => LegacyParent(c) is { } p && p.Author == au && p.Idx == ix).ToList()) k.Remove();
                cm.Remove();
                if (list.Elements(Ns.P + "cm").Any()) WritePart(lp, list); else sp.DeletePart(lp);
                Dirty = true; return;
            }
        }
        throw new KeyNotFoundException("Commento non trovato.");
    }
}
