using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Core;

/// <summary>Un commento o una risposta. Id identifica la voce nella slide (GUID moderno, "autore:idx" classico).</summary>
public sealed record CommentInfo(string Author, string Text, string Created, string Id = "", bool IsReply = false, int Replies = 0);

/// <summary>Intestazione e piè di pagina: numero slide, data (fissa o automatica), piè di pagina.</summary>
public sealed record HeaderFooter(bool SlideNumber, bool Date, bool DateAuto, string DateText, bool Footer, string FooterText, bool SkipTitle = false);

/// <summary>Opzioni di Imposta presentazione (p:showPr). Mode: present, browse, kiosk. From/To: intervallo 1-based o nulli per tutte.</summary>
public sealed record ShowSettings(string Mode, bool Loop, bool Narration, bool Animation, bool UseTimings, int? From, int? To);

/// <summary>Commenti, intestazione/piè di pagina, impostazioni della presentazione.</summary>
public sealed partial class PptxDocument
{
    const string P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
    const string CommentAuthorName = "Autore";   // ponytail: autore fisso, mai il nome dell'utente di Windows → impostazione

    static XNamespace N188 => P188;

    static void WritePart(OpenXmlPart p, XElement x)
    {
        using var ms = new MemoryStream(); using (var w = System.Xml.XmlWriter.Create(ms, new System.Xml.XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false) })) new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), x).Save(w);
        ms.Position = 0; p.FeedData(ms);
    }

    // ---------- commenti ---------- (lettura, modifica, eliminazione: PptxDocument.Comments.cs)
    Dictionary<string, string> ModernAuthors() =>
        PresentationPart.GetPartsOfType<PowerPointAuthorsPart>().FirstOrDefault() is { } ap ? Xml(ap).Elements(N188 + "author").ToDictionary(a => (string?)a.Attribute("id") ?? "", a => (string?)a.Attribute("name") ?? "?") : new();

    Dictionary<string, string> LegacyAuthors() =>
        PresentationPart.CommentAuthorsPart is { } ap ? Xml(ap).Elements(Ns.P + "cmAuthor").ToDictionary(a => (string?)a.Attribute("id") ?? "", a => (string?)a.Attribute("name") ?? "?") : new();

    /// <summary>Aggiunge un commento alla slide. Moderno (come PowerPoint 365) di norma; se la slide ha già commenti classici si aggiunge un commento classico.</summary>
    public void AddComment(int slide, string text)
    {
        text = CleanCommentText(text);
        var sp = Slide(slide);
        var now = DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss.fff", CultureInfo.InvariantCulture);
        if (sp.SlideCommentsPart != null && !sp.GetPartsOfType<PowerPointCommentPart>().Any()) AddLegacyComment(sp, text, now); else AddModernComment(sp, slide, text, now);
        Dirty = true;
    }

    void AddLegacyComment(SlidePart sp, string text, string now)
    {
        var ap = PresentationPart.CommentAuthorsPart ?? PresentationPart.AddNewPart<CommentAuthorsPart>();
        XElement authors = ap.CommentAuthorList != null || ap.GetStream().Length > 0 ? Xml(ap) : new XElement(Ns.P + "cmAuthorLst", new XAttribute(XNamespace.Xmlns + "p", Ns.P.NamespaceName));
        var au = authors.Elements(Ns.P + "cmAuthor").FirstOrDefault(a => (string?)a.Attribute("name") == CommentAuthorName);
        if (au == null)
        {
            int nid = authors.Elements(Ns.P + "cmAuthor").Select(a => (int?)a.Attribute("id") ?? 0).DefaultIfEmpty(-1).Max() + 1;
            au = new XElement(Ns.P + "cmAuthor", new XAttribute("id", nid), new XAttribute("name", CommentAuthorName), new XAttribute("initials", "A"), new XAttribute("lastIdx", 0), new XAttribute("clrIdx", nid % 8));
            authors.Add(au);
        }
        int idx = (int?)au.Attribute("lastIdx") + 1 ?? 1; au.SetAttributeValue("lastIdx", idx);
        WritePart(ap, authors);
        var lp = sp.SlideCommentsPart!; var list = Xml(lp);
        int n = list.Elements(Ns.P + "cm").Count();
        list.Add(new XElement(Ns.P + "cm", new XAttribute("authorId", (string)au.Attribute("id")!), new XAttribute("dt", now), new XAttribute("idx", idx),
            new XElement(Ns.P + "pos", new XAttribute("x", 10 + n * 10), new XAttribute("y", 10 + n * 10)), new XElement(Ns.P + "text", text)));
        WritePart(lp, list);
    }

    void AddModernComment(SlidePart sp, int slide, string text, string now)
    {
        // autore
        var ap = PresentationPart.GetPartsOfType<PowerPointAuthorsPart>().FirstOrDefault();
        XElement authors = ap != null ? Xml(ap) : new XElement(N188 + "authorLst", new XAttribute(XNamespace.Xmlns + "a", Ns.A.NamespaceName), new XAttribute(XNamespace.Xmlns + "r", Ns.R.NamespaceName), new XAttribute(XNamespace.Xmlns + "p188", P188));
        var au = authors.Elements(N188 + "author").FirstOrDefault(a => (string?)a.Attribute("name") == CommentAuthorName);
        if (au == null)
        {
            au = new XElement(N188 + "author", new XAttribute("id", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"), new XAttribute("name", CommentAuthorName), new XAttribute("initials", "A"), new XAttribute("userId", CommentAuthorName), new XAttribute("providerId", "None"));
            authors.Add(au);
        }
        ap ??= PresentationPart.AddNewPart<PowerPointAuthorsPart>();
        WritePart(ap, authors);
        // id di creazione della slide (p14:creationId) richiesto dal marcatore del commento
        var creation = EnsureCreationId(sp);
        uint sldId = SlideIds!.Elements<P.SlideId>().ElementAt(slide).Id!.Value;
        var cp = sp.GetPartsOfType<PowerPointCommentPart>().FirstOrDefault();
        XElement list = cp != null ? Xml(cp) : new XElement(N188 + "cmLst", new XAttribute(XNamespace.Xmlns + "a", Ns.A.NamespaceName), new XAttribute(XNamespace.Xmlns + "r", Ns.R.NamespaceName), new XAttribute(XNamespace.Xmlns + "p188", P188));
        int n = list.Elements(N188 + "cm").Count();
        XNamespace pc = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
        var body = new XElement(N188 + "txBody", new XElement(Ns.A + "bodyPr"), new XElement(Ns.A + "lstStyle"),
            CommentParas(text));
        list.Add(new XElement(N188 + "cm", new XAttribute("id", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"), new XAttribute("authorId", (string)au.Attribute("id")!), new XAttribute("created", now),
            new XElement(pc + "sldMkLst", new XAttribute(XNamespace.Xmlns + "pc", pc.NamespaceName), new XElement(pc + "docMk"), new XElement(pc + "sldMk", new XAttribute("cId", creation), new XAttribute("sldId", sldId))),
            new XElement(N188 + "pos", new XAttribute("x", 1270000 + n * 254000), new XAttribute("y", 1270000 + n * 254000)), body));
        if (cp == null)
        {
            cp = sp.AddNewPart<PowerPointCommentPart>();
            var rid = sp.GetIdOfPart(cp);
            var ext = new P.SlideExtension($"<p:ext xmlns:p=\"{Ns.P.NamespaceName}\" uri=\"{{6950BFC3-D8DA-4A85-94F7-54DA5524770B}}\"><p188:commentRel xmlns:p188=\"{P188}\" xmlns:r=\"{Ns.R.NamespaceName}\" r:id=\"{rid}\"/></p:ext>");
            var sl = sp.Slide!;
            if (sl.GetFirstChild<P.SlideExtensionList>() is { } el) el.AppendChild(ext); else sl.AppendChild(new P.SlideExtensionList(ext));
        }
        WritePart(cp, list);
    }

    static string EnsureCreationId(SlidePart sp)
    {
        var cs = sp.Slide!.CommonSlideData!;
        var ex = cs.GetFirstChild<P.CommonSlideDataExtensionList>();
        var found = ex?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.CreationId>().FirstOrDefault();
        if (found?.Val != null) return found.Val.Value.ToString(CultureInfo.InvariantCulture);
        uint val = (uint)Random.Shared.NextInt64(1, uint.MaxValue);
        var ext = new P.CommonSlideDataExtension { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };
        ext.AppendChild(new DocumentFormat.OpenXml.Office2010.PowerPoint.CreationId { Val = val });
        if (ex == null) cs.AppendChild(new P.CommonSlideDataExtensionList(ext)); else ex.AppendChild(ext);
        return val.ToString(CultureInfo.InvariantCulture);
    }

    // ---------- intestazione e piè di pagina ----------
    static readonly (string Type, string Name)[] HfTypes = { ("dt", "Segnaposto data"), ("ftr", "Segnaposto piè di pagina"), ("sldNum", "Segnaposto numero diapositiva") };

    P.Shape? HfFind(SlidePart sp, string type)
    {
        foreach (var s in sp.Slide!.Descendants<P.Shape>())
        {
            var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
            if (ph?.Type == null) continue;
            var v = ph.Type.Value;
            if ((type == "dt" && v == P.PlaceholderValues.DateAndTime) || (type == "ftr" && v == P.PlaceholderValues.Footer) || (type == "sldNum" && v == P.PlaceholderValues.SlideNumber)) return s;
        }
        return null;
    }

    static P.PlaceholderShape? LayoutPh(SlideLayoutPart lp, string type)
    {
        foreach (var s in lp.SlideLayout!.Descendants<P.Shape>())
        {
            var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
            if (ph?.Type == null) continue;
            var v = ph.Type.Value;
            if ((type == "dt" && v == P.PlaceholderValues.DateAndTime) || (type == "ftr" && v == P.PlaceholderValues.Footer) || (type == "sldNum" && v == P.PlaceholderValues.SlideNumber)) return ph;
        }
        return null;
    }

    static string Esc(string s) => System.Security.SecurityElement.Escape(s);

    /// <summary>Stato letto dalla slide: numero, data (fissa o automatica), piè di pagina.</summary>
    public HeaderFooter GetHeaderFooter(int slide)
    {
        var sp = Slide(slide);
        var dt = HfFind(sp, "dt"); var ft = HfFind(sp, "ftr"); var nu = HfFind(sp, "sldNum");
        string Txt(P.Shape? s) => s?.TextBody == null ? "" : string.Concat(s.TextBody.Descendants<A.Text>().Select(t => t.Text));
        bool auto = dt?.TextBody?.Descendants<A.Field>().Any() == true;
        return new HeaderFooter(nu != null, dt != null, auto, auto || dt == null ? DateTime.Today.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : Txt(dt), ft != null, Txt(ft));
    }

    /// <summary>Applica a una slide o a tutte (slide = -1). Un segnaposto che il layout non prevede non viene creato (come PowerPoint).</summary>
    public void ApplyHeaderFooter(HeaderFooter hf, int slide = -1)
    {
        if (hf.FooterText.Length > 500 || hf.DateText.Length > 200) throw new ArgumentException("Testo troppo lungo.");
        for (int i = 0; i < Slides.Count; i++)
        {
            if (slide >= 0 && i != slide) continue;
            var sp = Slide(i); var tree = Tree(sp);
            var layout = sp.SlideLayoutPart;
            bool title = hf.SkipTitle && layout?.SlideLayout?.Type?.Value == P.SlideLayoutValues.Title;
            foreach (var (type, name) in HfTypes)
            {
                bool want = !title && type switch { "dt" => hf.Date, "ftr" => hf.Footer, _ => hf.SlideNumber };
                var ex = HfFind(sp, type);
                if (!want) { ex?.Remove(); continue; }
                var lph = layout == null ? null : LayoutPh(layout, type);
                if (ex == null && lph == null) continue;
                string guid = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
                string para = type switch
                {
                    "sldNum" => $"<a:p><a:fld id=\"{guid}\" type=\"slidenum\"><a:rPr lang=\"it-IT\"/><a:t>{i + 1}</a:t></a:fld><a:endParaRPr lang=\"it-IT\"/></a:p>",
                    "dt" when hf.DateAuto => $"<a:p><a:fld id=\"{guid}\" type=\"datetime1\"><a:rPr lang=\"it-IT\"/><a:t>{DateTime.Today.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}</a:t></a:fld><a:endParaRPr lang=\"it-IT\"/></a:p>",
                    "dt" => $"<a:p><a:r><a:rPr lang=\"it-IT\"/><a:t>{Esc(hf.DateText)}</a:t></a:r><a:endParaRPr lang=\"it-IT\"/></a:p>",
                    _ => $"<a:p><a:r><a:rPr lang=\"it-IT\"/><a:t>{Esc(hf.FooterText)}</a:t></a:r><a:endParaRPr lang=\"it-IT\"/></a:p>"
                };
                if (ex == null)
                {
                    int id = NextShapeId(sp);
                    string attrs = $" type=\"{type}\"" + (lph!.Size?.InnerText is { Length: > 0 } sz ? $" sz=\"{sz}\"" : "") + (lph.Index?.Value is { } ix ? $" idx=\"{ix}\"" : "");
                    ex = new P.Shape($"<p:sp xmlns:p=\"{Ns.P.NamespaceName}\" xmlns:a=\"{Ns.A.NamespaceName}\"><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"{name} {id - 1}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{para}</p:txBody></p:sp>");
                    tree.AppendChild(ex);
                }
                else
                {
                    ex.TextBody ??= new P.TextBody(new A.BodyProperties(), new A.ListStyle());
                    foreach (var p in ex.TextBody.Elements<A.Paragraph>().ToList()) p.Remove();
                    ex.TextBody.AppendChild(new A.Paragraph(para.Replace("<a:p>", $"<a:p xmlns:a=\"{Ns.A.NamespaceName}\">")));
                }
            }
        }
        Dirty = true;
    }

    // ---------- imposta presentazione ----------
    public ShowSettings GetShowSettings()
    {
        var sh = PresentationPart.PresentationPropertiesPart?.PresentationProperties?.GetFirstChild<P.ShowProperties>();
        if (sh == null) return new ShowSettings("present", false, true, true, true, null, null);
        string mode = sh.GetFirstChild<P.BrowseSlideMode>() != null ? "browse" : sh.GetFirstChild<P.KioskSlideMode>() != null ? "kiosk" : "present";
        var rg = sh.GetFirstChild<P.SlideRange>();
        return new ShowSettings(mode, sh.Loop?.Value ?? false, sh.ShowNarration?.Value ?? true, sh.ShowAnimation?.Value ?? true, sh.UseTimings?.Value ?? true, (int?)rg?.Start?.Value, (int?)rg?.End?.Value);
    }

    public void SetShowSettings(ShowSettings s)
    {
        if (s.Mode is not ("present" or "browse" or "kiosk")) throw new ArgumentException(s.Mode);
        int count = Slides.Count;
        if (s.From != null || s.To != null)
        {
            if (s.From is not { } f || s.To is not { } t || f < 1 || t < f || t > count) throw new ArgumentException($"Intervallo non valido: da 1 a {count}, con inizio non oltre la fine.");
        }
        var pp = PresentationPart.PresentationPropertiesPart ?? PresentationPart.AddNewPart<PresentationPropertiesPart>();
        pp.PresentationProperties ??= new P.PresentationProperties();
        var props = pp.PresentationProperties;
        var sh = props.GetFirstChild<P.ShowProperties>();
        if (sh == null)
        {
            sh = new P.ShowProperties();
            var after = props.ChildElements.LastOrDefault(c => c is P.HtmlPublishProperties or P.WebProperties or P.PrintingProperties);
            var before = props.ChildElements.FirstOrDefault(c => c is P.ColorMostRecentlyUsed or P.PresentationPropertiesExtensionList);
            if (after != null) after.InsertAfterSelf(sh); else if (before != null) before.InsertBeforeSelf(sh); else props.AppendChild(sh);
        }
        sh.Loop = s.Loop ? true : null; sh.ShowNarration = s.Narration ? null : false; sh.ShowAnimation = s.Animation ? null : false; sh.UseTimings = s.UseTimings ? null : false;
        foreach (var c in sh.ChildElements.Where(c => c is P.PresenterSlideMode or P.BrowseSlideMode or P.KioskSlideMode or P.SlideAll or P.SlideRange or P.CustomShowReference).ToList()) c.Remove();
        OpenXmlElement mode = s.Mode switch { "browse" => new P.BrowseSlideMode(), "kiosk" => new P.KioskSlideMode(), _ => new P.PresenterSlideMode() };
        OpenXmlElement range = s.From != null ? new P.SlideRange { Start = (uint)s.From.Value, End = (uint)s.To!.Value } : new P.SlideAll();
        sh.InsertAt(mode, 0); mode.InsertAfterSelf(range);
        Dirty = true;
    }
}
