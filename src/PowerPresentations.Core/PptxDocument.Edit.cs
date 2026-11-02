using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.Core;

public sealed record TransitionInfo(string Kind, string? Dir, string Speed);
public sealed record LayoutInfo(string Name, SlideLayoutPart Part);

/// <summary>Operazioni di modifica: toccano solo gli elementi indicati, tutto il resto del package resta intatto.</summary>
public sealed partial class PptxDocument
{
    const string NsDecl = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
    static long Emu(double pt) => (long)Math.Round(pt * 12700);

    SlidePart Slide(int i) => Slides.ElementAtOrDefault(i) ?? throw new ArgumentOutOfRangeException(nameof(i));

    public IReadOnlyList<LayoutInfo> Layouts =>
        PresentationPart.SlideMasterParts.SelectMany(m => m.SlideLayoutParts)
            .Select(l => new LayoutInfo(l.SlideLayout?.CommonSlideData?.Name?.Value ?? "Layout", l)).ToList();

    // ---------- slide ----------
    public bool IsHidden(int i) => Slide(i).Slide?.Show?.Value == false;

    public void SetHidden(int i, bool hidden) { Slide(i).Slide!.Show = hidden ? false : null; Dirty = true; }

    /// <summary>Sposta una slide: modifica solo sldIdLst; part, relazioni, note e commenti restano com'erano.</summary>
    public void MoveSlide(int from, int to)
    {
        var ids = SlideIds!.Elements<P.SlideId>().ToList();
        if (from < 0 || from >= ids.Count || to < 0 || to >= ids.Count) throw new ArgumentOutOfRangeException(nameof(to));
        if (from == to) return;
        var e = ids[from]; e.Remove();
        var rest = SlideIds.Elements<P.SlideId>().ToList();
        if (to >= rest.Count) SlideIds.AppendChild(e); else rest[to].InsertBeforeSelf(e);
        Dirty = true;
    }

    uint NextSlideId() => Math.Max(255, SlideIds?.Elements<P.SlideId>().Select(s => s.Id?.Value ?? 0).DefaultIfEmpty(255u).Max() ?? 255u) + 1;

    int InsertSlideId(SlidePart np, int afterIndex)
    {
        var lst = PresentationPart.Presentation!.SlideIdList ?? PresentationPart.Presentation.InsertAfter(new P.SlideIdList(), PresentationPart.Presentation.GetFirstChild<P.SlideMasterIdList>()!);
        var sid = new P.SlideId { Id = NextSlideId(), RelationshipId = PresentationPart.GetIdOfPart(np) };
        var cur = lst.Elements<P.SlideId>().ToList();
        if (afterIndex < 0 || afterIndex >= cur.Count - 0 && cur.Count == 0) lst.AppendChild(sid);
        else if (afterIndex >= cur.Count - 1) lst.AppendChild(sid);
        else cur[afterIndex].InsertAfterSelf(sid);
        Dirty = true;
        return lst.Elements<P.SlideId>().ToList().IndexOf(sid);
    }

    public int AddSlide(LayoutInfo layout, int afterIndex)
    {
        var np = PresentationPart.AddNewPart<SlidePart>();
        np.AddPart(layout.Part);
        var shapes = new System.Text.StringBuilder(); int id = 2, n = 1;
        var lx = Xml(layout.Part);
        foreach (var sp in lx.Descendants(Ns.P + "sp"))
        {
            var ph = sp.Element(Ns.P + "nvSpPr")?.Element(Ns.P + "nvPr")?.Element(Ns.P + "ph"); if (ph == null) continue;
            var type = (string?)ph.Attribute("type");
            if (type is "dt" or "ftr" or "sldNum") continue;
            string attrs = (type != null ? $" type=\"{type}\"" : "") + ((string?)ph.Attribute("idx") is { } ix ? $" idx=\"{ix}\"" : "") + ((string?)ph.Attribute("sz") is { } sz ? $" sz=\"{sz}\"" : "");
            string tx = type is "pic" or "media" ? "" : "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"it-IT\"/></a:p></p:txBody>";
            shapes.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{++id - 1}\" name=\"Segnaposto {n++}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr><p:spPr/>{tx}</p:sp>");
        }
        np.Slide = new P.Slide($"<p:sld {NsDecl}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
        return InsertSlideId(np, afterIndex);
    }

    /// <summary>Duplica una slide copiando XML e relazioni (immagini condivise). Le note si copiano come testo; i commenti no (ponytail).</summary>
    public int DuplicateSlide(int i)
    {
        var src = Slide(i);
        var np = PresentationPart.AddNewPart<SlidePart>();
        // copia dal DOM corrente (le modifiche non ancora salvate sono nel DOM, non nello stream della part)
        using (var s = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(src.Slide!.OuterXml))) np.FeedData(s);
        foreach (var pr in src.Parts)
        {
            if (pr.OpenXmlPart is NotesSlidePart or SlideCommentsPart or PowerPointCommentPart) continue;
            if (pr.OpenXmlPart is ChartPart or DiagramDataPart or DiagramLayoutDefinitionPart or DiagramStylePart or DiagramColorsPart or DiagramPersistLayoutPart) CloneInto(np, pr.OpenXmlPart, pr.RelationshipId);
            else np.AddPart(pr.OpenXmlPart, pr.RelationshipId);
        }
        foreach (var h in src.HyperlinkRelationships) np.AddHyperlinkRelationship(h.Uri, h.IsExternal, h.Id);
        foreach (var x in src.ExternalRelationships) np.AddExternalRelationship(x.RelationshipType, x.Uri, x.Id);
        var idx = InsertSlideId(np, i);
        var notes = GetNotes(i); if (notes.Length > 0) SetNotes(idx, notes);
        return idx;
    }

    /// <summary>Copia profonda di una part (grafico, SmartArt) con i suoi figli: PowerPoint non accetta la stessa part di grafico condivisa da due slide. Le immagini restano condivise.</summary>
    static void CloneInto(OpenXmlPartContainer parent, OpenXmlPart src, string rid)
    {
        var ext = System.IO.Path.GetExtension(src.Uri.OriginalString); if (ext.Length == 0) ext = ".xml";
        var copy = parent.AddExtendedPart(src.RelationshipType, src.ContentType, ext, rid);
        using (var s = src.GetStream(FileMode.Open, FileAccess.Read)) copy.FeedData(s);
        foreach (var c in src.Parts)
        {
            if (c.OpenXmlPart is ImagePart) copy.AddPart(c.OpenXmlPart, c.RelationshipId); else CloneInto(copy, c.OpenXmlPart, c.RelationshipId);
        }
        foreach (var h in src.ExternalRelationships) copy.AddExternalRelationship(h.RelationshipType, h.Uri, h.Id);
    }

    public void DeleteSlide(int i)
    {
        var sp = Slide(i);
        var rid = PresentationPart.GetIdOfPart(sp);
        var sid = SlideIds!.Elements<P.SlideId>().First(s => s.RelationshipId == rid);
        var idVal = sid.Id!.Value;
        sid.Remove();
        // riferimenti nelle sezioni (p14:sldId) e nelle presentazioni personalizzate (p:sldId r:id)
        foreach (var e in PresentationPart.Presentation!.Descendants<OpenXmlElement>().Where(e => e.LocalName == "sldId" && e.Parent != SlideIds).ToList())
            if (e.GetAttributes().Any(a => a.LocalName == "id" && a.Value == idVal.ToString()) || e.GetAttributes().Any(a => a.LocalName == "id" && a.NamespaceUri.EndsWith("/relationships") && a.Value == rid))
                e.Remove();
        if (sp.NotesSlidePart is { } np) sp.DeletePart(np);
        PresentationPart.DeletePart(sp);
        Dirty = true;
    }

    // ---------- note ----------
    public string GetNotes(int i)
    {
        var ns = Slide(i).NotesSlidePart; if (ns == null) return "";
        var body = ns.NotesSlide?.Descendants<P.Shape>().FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Body);
        return body?.TextBody == null ? "" : string.Join("\n", body.TextBody.Elements<A.Paragraph>().Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text))));
    }

    NotesMasterPart EnsureNotesMaster()
    {
        if (PresentationPart.NotesMasterPart is { } nm) return nm;
        nm = PresentationPart.AddNewPart<NotesMasterPart>();
        var theme = nm.AddNewPart<ThemePart>();
        var src = PresentationPart.SlideMasterParts.First().ThemePart!;
        using (var s = src.GetStream()) theme.FeedData(s);
        nm.NotesMaster = new P.NotesMaster($"<p:notesMaster {NsDecl}><p:cSld><p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>" +
            "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Immagine diapositiva\"/><p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"sldImg\" idx=\"2\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"685800\" y=\"1143000\"/><a:ext cx=\"5486400\" cy=\"3086100\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"12700\"><a:solidFill><a:prstClr val=\"black\"/></a:solidFill></a:ln></p:spPr></p:sp>" +
            "<p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Note\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"quarter\" idx=\"3\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"685800\" y=\"4400550\"/><a:ext cx=\"5486400\" cy=\"3600450\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/><a:lstStyle/><a:p><a:pPr lvl=\"0\"/><a:r><a:rPr lang=\"it-IT\"/><a:t>Modifica gli stili del testo dello schema</a:t></a:r></a:p></p:txBody></p:sp>" +
            "</p:spTree></p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/><p:notesStyle><a:lvl1pPr marL=\"0\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>");
        var pres = PresentationPart.Presentation!;
        var lst = new P.NotesMasterIdList(new P.NotesMasterId { Id = PresentationPart.GetIdOfPart(nm) });
        pres.InsertAfter(lst, pres.GetFirstChild<P.SlideMasterIdList>()!);
        return nm;
    }

    public void SetNotes(int i, string text)
    {
        var sp = Slide(i);
        var ns = sp.NotesSlidePart;
        if (ns == null)
        {
            var nm = EnsureNotesMaster();
            ns = sp.AddNewPart<NotesSlidePart>();
            ns.AddPart(nm); ns.AddPart(sp);
            ns.NotesSlide = new P.NotesSlide($"<p:notes {NsDecl}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>" +
                "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Immagine diapositiva 1\"/><p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"sldImg\"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>" +
                "<p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Segnaposto note 2\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"it-IT\"/></a:p></p:txBody></p:sp>" +
                "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>");
        }
        var body = ns.NotesSlide!.Descendants<P.Shape>().FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Body)
            ?? throw new InvalidOperationException("La pagina delle note non ha il segnaposto del testo.");
        FillText(body, text);
        Dirty = true;
    }

    // ---------- testo ----------
    static P.Shape? FindShape(SlidePart sp, int id) =>
        sp.Slide!.Descendants<P.Shape>().FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)id);

    public string GetShapeText(int slide, int id)
    {
        var sh = FindShape(Slide(slide), id);
        return sh?.TextBody == null ? "" : string.Join("\n", sh.TextBody.Elements<A.Paragraph>().Select(p => string.Concat(p.Elements().Select(e => e switch { A.Run r => r.Text?.Text ?? "", A.Field f => f.Text?.Text ?? "", A.Break => "\v", _ => "" }))));
    }

    /// <summary>Sostituisce il testo mantenendo proprietà di paragrafo e formato della prima run di ogni paragrafo (ponytail: le run con formati diversi nello stesso paragrafo si fondono nel formato della prima).</summary>
    public void SetShapeText(int slide, int id, string text)
    {
        var sh = FindShape(Slide(slide), id) ?? throw new KeyNotFoundException($"Forma {id} non trovata.");
        FillText(sh, text); Dirty = true;
        Refit(slide, id, sh);
    }

    /// <summary>Dopo una modifica del testo ricalcola il rimpicciolimento automatico (normAutofit) e lo memorizza come fa PowerPoint (fontScale, lnSpcReduction).</summary>
    void Refit(int slide, int id, P.Shape sh)
    {
        var fits = new SlideRenderer(this).Render(slide, new RenderOptions { RecomputeAutofit = true }).Fits;
        if (!fits.TryGetValue(id, out var f)) return;
        var bp = sh.TextBody!.BodyProperties ??= new A.BodyProperties();
        var na = bp.GetFirstChild<A.NormalAutoFit>();
        if (na == null)
        {
            na = new A.NormalAutoFit();
            var warp = bp.GetFirstChild<A.PresetTextWarp>();
            if (warp != null) warp.InsertAfterSelf(na); else bp.InsertAt(na, 0);
        }
        na.FontScale = f.Scale < 1 ? (int)Math.Round(f.Scale * 100000) : null;
        na.LineSpaceReduction = f.Reduction > 0 && f.Scale < 1 ? (int)Math.Round(f.Reduction * 100000) : null;
    }

    static void FillText(P.Shape sh, string text)
    {
        sh.TextBody ??= new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph());
        FillBody(sh.TextBody, text);
    }

    static void FillBody(OpenXmlCompositeElement tb, string text)
    {
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var paras = tb.Elements<A.Paragraph>().ToList();
        if (paras.Count == 0) { var np = new A.Paragraph(); tb.AppendChild(np); paras.Add(np); }
        for (int i = 0; i < lines.Length; i++)
        {
            A.Paragraph p;
            if (i < paras.Count) p = paras[i];
            else { p = (A.Paragraph)paras[^1].CloneNode(true); tb.AppendChild(p); }
            var first = p.Elements<A.Run>().FirstOrDefault();
            var rpr = (A.RunProperties?)(first?.RunProperties ?? (p.GetFirstChild<A.EndParagraphRunProperties>() is { } e ? ToRunProps(e) : null))?.CloneNode(true);
            foreach (var c in p.Elements().Where(c => c is A.Run or A.Field or A.Break).ToList()) c.Remove();
            if (lines[i].Length > 0)
            {
                var run = new A.Run(); if (rpr != null) run.AppendChild(rpr); else run.AppendChild(new A.RunProperties { Language = "it-IT" });
                run.AppendChild(new A.Text(lines[i]));
                var end = p.GetFirstChild<A.EndParagraphRunProperties>();
                if (end != null) end.InsertBeforeSelf(run); else p.AppendChild(run);
            }
        }
        for (int i = lines.Length; i < paras.Count; i++) paras[i].Remove();
    }

    static A.RunProperties ToRunProps(A.EndParagraphRunProperties e)
    {
        var r = new A.RunProperties();
        foreach (var a in e.GetAttributes()) r.SetAttribute(a);
        foreach (var c in e.ChildElements) r.AppendChild(c.CloneNode(true));
        return r;
    }

    /// <summary>Formato carattere su tutto il testo della forma.</summary>
    public void SetFormat(int slide, int id, bool? bold = null, bool? italic = null, bool? underline = null, double? sizePt = null, string? colorHex = null)
    {
        var bodies = TextBodies(slide, id);
        void Apply(OpenXmlElement rp)
        {
            if (bold != null) rp.SetAttribute(new OpenXmlAttribute("b", "", bold.Value ? "1" : "0"));
            if (italic != null) rp.SetAttribute(new OpenXmlAttribute("i", "", italic.Value ? "1" : "0"));
            if (underline != null) rp.SetAttribute(new OpenXmlAttribute("u", "", underline.Value ? "sng" : "none"));
            if (sizePt != null) rp.SetAttribute(new OpenXmlAttribute("sz", "", ((int)Math.Round(sizePt.Value * 100)).ToString()));
            if (colorHex != null)
            {
                foreach (var f in rp.Elements<A.SolidFill>().ToList()) f.Remove();
                var fill = new A.SolidFill(new A.RgbColorModelHex { Val = colorHex });
                var ln = rp.GetFirstChild<A.Outline>();
                if (ln != null) ln.InsertAfterSelf(fill); else rp.InsertAt(fill, 0);
            }
        }
        foreach (var tb in bodies)
        {
            foreach (var r in tb.Descendants<A.Run>()) { r.RunProperties ??= new A.RunProperties { Language = "it-IT" }; Apply(r.RunProperties); }
            foreach (var e in tb.Descendants<A.EndParagraphRunProperties>()) Apply(e);
        }
        if (bodies.Count > 0) Dirty = true;
    }

    /// <summary>Corpi di testo di una forma o, per una tabella, di tutte le celle. Lista vuota se l'oggetto non ha testo (immagine); KeyNotFound se l'id non esiste.</summary>
    public List<OpenXmlCompositeElement> TextBodies(int slide, int id)
    {
        var sp = Slide(slide);
        if (FindShape(sp, id) is { } sh) return sh.TextBody == null ? new() : new() { sh.TextBody };
        var gf = sp.Slide!.Descendants<P.GraphicFrame>().FirstOrDefault(g => g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)id);
        if (gf != null) return gf.Descendants<A.TextBody>().Cast<OpenXmlCompositeElement>().ToList();
        if (sp.Slide.Descendants<P.NonVisualDrawingProperties>().Any(n => n.Id?.Value == (uint)id)) return new();
        throw new KeyNotFoundException($"Forma {id} non trovata.");
    }

    /// <summary>Valori scritti nel file sulla prima run: (sottolineato, dimensione in punti). Nulla se non esplicito (ereditato).</summary>
    public (bool Underline, double? SizePt) GetRunState(int slide, int id)
    {
        var r = TextBodies(slide, id).SelectMany(t => t.Descendants<A.Run>()).FirstOrDefault()?.RunProperties;
        var u = r?.Underline?.Value; var sz = r?.FontSize?.Value;
        return (u != null && u != A.TextUnderlineValues.None, sz == null ? null : sz.Value / 100.0);
    }

    // ---------- inserimento oggetti ----------
    int NextShapeId(SlidePart sp) => (int)Math.Max(1, sp.Slide!.Descendants<P.NonVisualDrawingProperties>().Select(n => n.Id?.Value ?? 0).DefaultIfEmpty(1u).Max()) + 1;
    P.ShapeTree Tree(SlidePart sp) => sp.Slide!.CommonSlideData!.ShapeTree!;

    public int AddTextBox(int slide, double x, double y, double w, double h, string text = "")
    {
        var sp = Slide(slide); int id = NextShapeId(sp);
        var shape = new P.Shape($"<p:sp {NsDecl}><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"CasellaDiTesto {id}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{Emu(x)}\" y=\"{Emu(y)}\"/><a:ext cx=\"{Emu(w)}\" cy=\"{Emu(h)}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap=\"square\" rtlCol=\"0\"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang=\"it-IT\"/></a:p></p:txBody></p:sp>");
        Tree(sp).AppendChild(shape);
        if (text.Length > 0) FillText(shape, text);
        Dirty = true; return id;
    }

    public int AddShape(int slide, string prst, double x, double y, double w, double h, string text = "")
    {
        var sp = Slide(slide); int id = NextShapeId(sp);
        var shape = new P.Shape($"<p:sp {NsDecl}><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"Forma {id}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{Emu(x)}\" y=\"{Emu(y)}\"/><a:ext cx=\"{Emu(w)}\" cy=\"{Emu(h)}\"/></a:xfrm><a:prstGeom prst=\"{prst}\"><a:avLst/></a:prstGeom></p:spPr>" +
            "<p:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></p:style>" +
            "<p:txBody><a:bodyPr rtlCol=\"0\" anchor=\"ctr\"/><a:lstStyle/><a:p><a:pPr algn=\"ctr\"/><a:endParaRPr lang=\"it-IT\"/></a:p></p:txBody></p:sp>");
        Tree(sp).AppendChild(shape);
        if (text.Length > 0) FillText(shape, text);
        Dirty = true; return id;
    }

    static string? ImageType(byte[] d) =>
        d.Length > 8 && d[0] == 0x89 && d[1] == 0x50 ? "image/png" : d.Length > 3 && d[0] == 0xFF && d[1] == 0xD8 ? "image/jpeg" :
        d.Length > 3 && d[0] == 'G' && d[1] == 'I' && d[2] == 'F' ? "image/gif" : d.Length > 2 && d[0] == 'B' && d[1] == 'M' ? "image/bmp" : null;

    public int AddPicture(int slide, byte[] data, double x, double y, double w, double h, string description = "Immagine")
    {
        var type = ImageType(data) ?? throw new PptxFormatException("Formato immagine non supportato (PNG, JPEG, GIF, BMP).");
        var sp = Slide(slide); int id = NextShapeId(sp);
        var ip = sp.AddImagePart(type);
        using (var ms = new MemoryStream(data)) ip.FeedData(ms);
        var rid = sp.GetIdOfPart(ip);
        var pic = new P.Picture($"<p:pic {NsDecl}><p:nvPicPr><p:cNvPr id=\"{id}\" name=\"Immagine {id}\" descr=\"{System.Security.SecurityElement.Escape(description)}\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"{Emu(x)}\" y=\"{Emu(y)}\"/><a:ext cx=\"{Emu(w)}\" cy=\"{Emu(h)}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>");
        Tree(sp).AppendChild(pic);
        Dirty = true; return id;
    }

    public int AddTable(int slide, int rows, int cols, double x, double y, double w, double rowH = 29.2)
    {
        if (rows < 1 || cols < 1 || rows > 100 || cols > 50) throw new ArgumentOutOfRangeException(nameof(rows));
        var sp = Slide(slide); int id = NextShapeId(sp);
        long cw = Emu(w) / cols, rh = Emu(rowH);
        var sb = new System.Text.StringBuilder();
        for (int r = 0; r < rows; r++)
        {
            sb.Append($"<a:tr h=\"{rh}\">");
            for (int c = 0; c < cols; c++) sb.Append("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"it-IT\"/></a:p></a:txBody><a:tcPr/></a:tc>");
            sb.Append("</a:tr>");
        }
        var grid = string.Concat(Enumerable.Repeat($"<a:gridCol w=\"{cw}\"/>", cols));
        var gf = new P.GraphicFrame($"<p:graphicFrame {NsDecl}><p:nvGraphicFramePr><p:cNvPr id=\"{id}\" name=\"Tabella {id}\"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"{Emu(x)}\" y=\"{Emu(y)}\"/><a:ext cx=\"{cw * cols}\" cy=\"{rh * rows}\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblPr firstRow=\"1\" bandRow=\"1\"><a:tableStyleId>{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}</a:tableStyleId></a:tblPr><a:tblGrid>{grid}</a:tblGrid>{sb}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
        Tree(sp).AppendChild(gf);
        Dirty = true; return id;
    }

    public void SetCellText(int slide, int frameId, int row, int col, string text)
    {
        var tc = Cell(slide, frameId, row, col);
        var tb = tc.TextBody ?? throw new InvalidOperationException("Cella senza testo.");
        FillBody(tb, text); Dirty = true;
    }

    public string GetCellText(int slide, int frameId, int row, int col)
    {
        var tc = Cell(slide, frameId, row, col);
        return tc.TextBody == null ? "" : string.Join("\n", tc.TextBody.Elements<A.Paragraph>().Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text))));
    }

    A.TableCell Cell(int slide, int frameId, int row, int col)
    {
        var gf = Slide(slide).Slide!.Descendants<P.GraphicFrame>().FirstOrDefault(g => g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)frameId) ?? throw new KeyNotFoundException();
        return gf.Descendants<A.TableRow>().ElementAt(row).Elements<A.TableCell>().ElementAt(col);
    }

    static readonly string[] TableFlags = { "firstRow", "lastRow", "firstCol", "lastCol", "bandRow", "bandCol" };

    public bool GetTableFlag(int slide, int frameId, string flag)
    {
        if (!TableFlags.Contains(flag)) throw new ArgumentException(flag);
        var pr = Cell(slide, frameId, 0, 0).Ancestors<A.Table>().First().TableProperties;
        return pr?.GetAttributes().FirstOrDefault(a => a.LocalName == flag).Value is "1" or "true";
    }

    /// <summary>Opzioni stile tabella (riga intestazione, righe alternate...).</summary>
    public void SetTableFlag(int slide, int frameId, string flag, bool on)
    {
        if (!TableFlags.Contains(flag)) throw new ArgumentException(flag);
        var t = Cell(slide, frameId, 0, 0).Ancestors<A.Table>().First();
        t.TableProperties ??= new A.TableProperties();
        t.TableProperties.SetAttribute(new OpenXmlAttribute(flag, "", on ? "1" : "0")); Dirty = true;
    }

    /// <summary>Allineamento dei paragrafi: l, ctr, r, just.</summary>
    public void SetAlignment(int slide, int id, string algn)
    {
        foreach (var p in TextBodies(slide, id).SelectMany(t => t.Elements<A.Paragraph>()))
        {
            p.ParagraphProperties ??= new A.ParagraphProperties();
            p.ParagraphProperties.SetAttribute(new OpenXmlAttribute("algn", "", algn));
        }
        Dirty = true;
    }

    public void DeleteShape(int slide, int id)
    {
        var sp = Slide(slide);
        var el = sp.Slide!.CommonSlideData!.ShapeTree!.Elements().FirstOrDefault(e => e.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value == (uint)id);
        el?.Remove(); Dirty = true;
    }

    /// <summary>Sposta/ridimensiona (punti). Se la forma eredita la posizione da un segnaposto, crea l'xfrm esplicito.</summary>
    public void SetShapeBox(int slide, int id, double x, double y, double w, double h)
    {
        var sp = Slide(slide);
        var el = sp.Slide!.CommonSlideData!.ShapeTree!.Elements().FirstOrDefault(e => e.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value == (uint)id) ?? throw new KeyNotFoundException();
        var off = new A.Offset { X = Emu(x), Y = Emu(y) }; var ext = new A.Extents { Cx = Emu(w), Cy = Emu(h) };
        switch (el)
        {
            case P.Shape s:
                {
                    s.ShapeProperties ??= new P.ShapeProperties();
                    var xf = s.ShapeProperties.Transform2D;
                    if (xf == null) { xf = new A.Transform2D(); s.ShapeProperties.InsertAt(xf, 0); }
                    xf.Offset = off; xf.Extents = ext; break;
                }
            case P.Picture pc:
                {
                    var xf = pc.ShapeProperties!.Transform2D ?? (pc.ShapeProperties.Transform2D = new A.Transform2D());
                    xf.Offset = off; xf.Extents = ext; break;
                }
            case P.GraphicFrame g:
                g.Transform ??= new P.Transform(); g.Transform.Offset = off; g.Transform.Extents = ext; break;
            default: throw new NotSupportedException("Tipo di forma non spostabile.");
        }
        Dirty = true;
    }

    // ---------- transizioni ----------
    public TransitionInfo GetTransition(int i)
    {
        var x = Xml(Slide(i));
        var t = x.Descendants().FirstOrDefault(e => e.Name.LocalName == "transition");
        if (t == null) return new TransitionInfo("none", null, "med");
        var k = t.Elements().FirstOrDefault(e => e.Name.LocalName != "sndAc" && e.Name.LocalName != "extLst");
        return new TransitionInfo(k?.Name.LocalName ?? "cut", (string?)k?.Attribute("dir"), (string?)t.Attribute("spd") ?? "med");
    }

    /// <param name="kind">none, fade, push, wipe, cover, pull, split, randomBar, cut</param>
    public void SetTransition(int i, string kind, string? dir = null, string speed = "med")
    {
        var sl = Slide(i).Slide!;
        // PowerPoint scrive la transizione in mc:AlternateContent con p14:dur (durata in ms): la durata esistente si conserva se la velocità non cambia
        string? oldDur = null, oldSpd = null;
        foreach (var ac in sl.Elements<AlternateContent>().Where(a => a.OuterXml.Contains("transition")).ToList())
        {
            var m = System.Text.RegularExpressions.Regex.Match(ac.OuterXml, "p14:dur=\"(\\d+)\"");
            if (m.Success) oldDur = m.Groups[1].Value;
            oldSpd = System.Text.RegularExpressions.Regex.Match(ac.OuterXml, "<p:transition[^>]* spd=\"(\\w+)\"").Groups[1].Value;
            ac.Remove();
        }
        foreach (var t in sl.Elements<P.Transition>().ToList()) t.Remove();
        if (kind != "none")
        {
            string inner = dir != null && kind is "push" or "wipe" or "cover" or "pull" or "split" or "randomBar" ? $"<p:{kind} dir=\"{dir}\"/>" : $"<p:{kind}/>";
            OpenXmlElement tr;
            if (oldDur != null)
            {
                var dur = oldSpd == speed ? oldDur : speed switch { "fast" => "500", "slow" => "2000", _ => "1000" };
                tr = new AlternateContent($"<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" {NsDecl}><mc:Choice xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" Requires=\"p14\"><p:transition spd=\"{speed}\" p14:dur=\"{dur}\">{inner}</p:transition></mc:Choice><mc:Fallback><p:transition spd=\"{speed}\">{inner}</p:transition></mc:Fallback></mc:AlternateContent>");
            }
            else tr = new P.Transition($"<p:transition {NsDecl} spd=\"{speed}\">{inner}</p:transition>");
            var anchor = (OpenXmlElement?)sl.GetFirstChild<P.ColorMapOverride>() ?? sl.CommonSlideData;
            anchor!.InsertAfterSelf(tr);
        }
        Dirty = true;
    }
}
