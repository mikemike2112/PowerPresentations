using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPresentations.Core;

public sealed record CustomShowInfo(string Name, IReadOnlyList<int> Slides);

/// <summary>Presentazioni personalizzate: elenco ordinato di slide con un nome (p:custShowLst in presentation.xml).</summary>
public sealed partial class PptxDocument
{
    P.CustomShowList? ShowList => PresentationPart.Presentation!.GetFirstChild<P.CustomShowList>();

    public IReadOnlyList<CustomShowInfo> CustomShows
    {
        get
        {
            var rid2idx = Slides.Select((s, i) => (rid: PresentationPart.GetIdOfPart(s), i)).ToDictionary(x => x.rid, x => x.i);
            return (ShowList?.Elements<P.CustomShow>() ?? Enumerable.Empty<P.CustomShow>())
                .Select(c => new CustomShowInfo(c.Name?.Value ?? "", c.SlideList?.Elements<P.SlideListEntry>().Select(e => e.Id?.Value is { } r && rid2idx.TryGetValue(r, out var i) ? i : -1).Where(i => i >= 0).ToList() ?? new List<int>())).ToList();
        }
    }

    /// <summary>Crea o sostituisce (stesso nome) una presentazione personalizzata con le slide indicate, nell'ordine dato.</summary>
    public void SetCustomShow(string name, IEnumerable<int> slides)
    {
        name = name.Trim();
        if (name.Length == 0 || name.Length > 100) throw new ArgumentException("Il nome deve avere da 1 a 100 caratteri.");
        var list = slides.ToList();
        if (list.Count == 0) throw new ArgumentException("Scegli almeno una diapositiva.");
        if (list.Any(i => i < 0 || i >= Slides.Count)) throw new ArgumentOutOfRangeException(nameof(slides));
        var pres = PresentationPart.Presentation!;
        var lst = ShowList;
        if (lst == null)
        {
            lst = new P.CustomShowList();
            var after = pres.ChildElements.LastOrDefault(c => c.LocalName is "sldIdLst" or "sldSz" or "notesSz" or "smartTags" or "embeddedFontLst");
            var before = pres.ChildElements.FirstOrDefault(c => c.LocalName is "photoAlbum" or "custDataLst" or "kinsoku" or "defaultTextStyle" or "modifyVerifier" or "extLst");
            if (after != null) after.InsertAfterSelf(lst); else if (before != null) before.InsertBeforeSelf(lst); else pres.AppendChild(lst);
        }
        var existing = lst.Elements<P.CustomShow>().FirstOrDefault(c => string.Equals(c.Name?.Value, name, StringComparison.CurrentCultureIgnoreCase));
        uint id = existing?.Id?.Value ?? (lst.Elements<P.CustomShow>().Any() ? lst.Elements<P.CustomShow>().Max(c => c.Id?.Value ?? 0) + 1 : 0u);
        var entries = list.Select(i => new P.SlideListEntry { Id = PresentationPart.GetIdOfPart(Slide(i)) });
        var cs = new P.CustomShow(new P.SlideList(entries)) { Name = name, Id = id };
        if (existing != null) existing.InsertBeforeSelf(cs); else lst.AppendChild(cs);
        existing?.Remove();
        Dirty = true;
    }

    public void DeleteCustomShow(string name)
    {
        var lst = ShowList; var c = lst?.Elements<P.CustomShow>().FirstOrDefault(x => string.Equals(x.Name?.Value, name, StringComparison.CurrentCultureIgnoreCase)) ?? throw new KeyNotFoundException($"Presentazione personalizzata «{name}» non trovata.");
        c.Remove(); if (!lst!.HasChildren) lst.Remove();
        Dirty = true;
    }
}
