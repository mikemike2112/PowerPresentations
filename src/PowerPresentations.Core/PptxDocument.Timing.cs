using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPresentations.Core;

/// <summary>Intervalli di avanzamento automatico (Prova intervalli): attributo advTm di p:transition.</summary>
public sealed partial class PptxDocument
{
    /// <summary>Millisecondi dopo cui la slide avanza da sola, o nulla.</summary>
    public int? GetAdvanceTime(int i)
    {
        var t = Slide(i).Slide!.Descendants<P.Transition>().FirstOrDefault();
        return t?.AdvanceAfterTime?.Value is { } s && int.TryParse(s, out var ms) && ms >= 0 ? ms : null;
    }

    /// <summary>Imposta (o con null toglie) l'avanzamento automatico, mantenendo l'effetto di transizione. Scrive sia in Choice sia in Fallback di mc:AlternateContent.</summary>
    public void SetAdvanceTime(int i, int? ms)
    {
        if (ms is < 0 or > 86_400_000) throw new ArgumentOutOfRangeException(nameof(ms));
        var sl = Slide(i).Slide!;
        var all = sl.Descendants<P.Transition>().ToList();
        if (all.Count == 0)
        {
            if (ms == null) return;
            var tr = new P.Transition(); tr.AdvanceAfterTime = ms.Value.ToString(); all.Add(tr);
            ((OpenXmlElement?)sl.GetFirstChild<P.ColorMapOverride>() ?? sl.CommonSlideData)!.InsertAfterSelf(tr);
        }
        foreach (var t in all) { if (ms == null) t.AdvanceAfterTime = null; else t.AdvanceAfterTime = ms.Value.ToString(); }
        // transizione rimasta vuota (nessun effetto, nessun attributo): si toglie
        foreach (var t in all.Where(IsEmptyTransition).ToList()) t.Remove();
        Dirty = true;
    }

    static bool IsEmptyTransition(P.Transition t) => !t.HasChildren && System.Text.RegularExpressions.Regex.IsMatch(t.OuterXml, @"^<p:transition(\s+xmlns:\w+=""[^""]*"")*\s*/>$");
}
