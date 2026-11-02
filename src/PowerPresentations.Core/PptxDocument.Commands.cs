using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPresentations.Core;

/// <summary>Comandi del ribbon: elenchi, transizione su tutte le slide.</summary>
public sealed partial class PptxDocument
{
    // ---------- elenchi ----------
    static readonly string[] BulletTags = { "buNone", "buAutoNum", "buChar", "buBlip", "buFont", "buFontTx" };

    /// <summary>Stato dell'elenco scritto nel file: none, bullet, number, mixed; inherit se nessun paragrafo ha un elenco esplicito.</summary>
    public string GetBullet(int slide, int id)
    {
        var kinds = TextBodies(slide, id).SelectMany(t => t.Elements<A.Paragraph>()).Select(p =>
        {
            var pr = p.ParagraphProperties;
            if (pr == null) return "inherit";
            if (pr.GetFirstChild<A.NoBullet>() != null) return "none";
            if (pr.GetFirstChild<A.AutoNumberedBullet>() != null) return "number";
            if (pr.GetFirstChild<A.CharacterBullet>() != null) return "bullet";
            return "inherit";
        }).Distinct().ToList();
        return kinds.Count == 0 ? "none" : kinds.Count == 1 ? kinds[0] : "mixed";
    }

    /// <summary>Imposta l'elenco su tutti i paragrafi: none, bullet, number. Ordine dei figli di a:pPr rispettato (bu* dopo spacing, prima di tabLst/defRPr).</summary>
    public void SetBullet(int slide, int id, string kind)
    {
        if (kind is not ("none" or "bullet" or "number")) throw new ArgumentException(kind);
        var paras = TextBodies(slide, id).SelectMany(t => t.Elements<A.Paragraph>()).ToList();
        foreach (var p in paras)
        {
            var pr = p.ParagraphProperties ??= new A.ParagraphProperties();
            foreach (var c in pr.ChildElements.Where(c => BulletTags.Contains(c.LocalName)).ToList()) c.Remove();
            OpenXmlElement[] add = kind switch
            {
                "none" => new OpenXmlElement[] { new A.NoBullet() },
                "bullet" => new OpenXmlElement[] { new A.BulletFont { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 }, new A.CharacterBullet { Char = "•" } },
                _ => new OpenXmlElement[] { new A.AutoNumberedBullet { Type = A.TextAutoNumberSchemeValues.ArabicPeriod } }
            };
            var anchor = pr.ChildElements.FirstOrDefault(c => c is A.TabStopList or A.DefaultRunProperties or A.ExtensionList);
            foreach (var e in add) { if (anchor != null) anchor.InsertBeforeSelf(e); else pr.AppendChild(e); }
            if (kind == "none") { pr.LeftMargin = 0; pr.Indent = 0; }
            else { if (pr.LeftMargin == null || pr.LeftMargin.Value == 0) pr.LeftMargin = 342900; if (pr.Indent == null || pr.Indent.Value >= 0) pr.Indent = -342900; }
        }
        if (paras.Count > 0) Dirty = true;
    }

    // ---------- transizione su tutte le slide ----------
    static bool IsTransitionNode(OpenXmlElement e) => e is P.Transition || (e is AlternateContent ac && ac.OuterXml.Contains("<p:transition"));

    /// <summary>Copia la transizione della slide (anche la forma mc:AlternateContent con durata p14) su tutte le altre; senza transizione le toglie.</summary>
    public void ApplyTransitionToAll(int from)
    {
        var src = Slide(from).Slide!;
        var node = src.ChildElements.FirstOrDefault(IsTransitionNode);
        for (int i = 0; i < Slides.Count; i++)
        {
            if (i == from) continue;
            var sl = Slide(i).Slide!;
            foreach (var t in sl.ChildElements.Where(IsTransitionNode).ToList()) t.Remove();
            if (node != null)
            {
                var anchor = (OpenXmlElement?)sl.GetFirstChild<P.ColorMapOverride>() ?? sl.CommonSlideData;
                anchor!.InsertAfterSelf(node.CloneNode(true));
            }
        }
        Dirty = true;
    }
}
