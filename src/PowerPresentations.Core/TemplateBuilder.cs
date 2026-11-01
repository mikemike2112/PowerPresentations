using System.Text;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPresentations.Core;

/// <summary>
/// Genera da codice il modello di "Nuova presentazione": tema "Prisma" (colori e font originali), schema diapositiva e 6 layout.
/// Nessun file derivato da Office: tutti i nomi, i colori e le misure sono scelti qui.
/// </summary>
public static class TemplateBuilder
{
    const string Ns = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
    const string Grp = "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>";

    static string Xfrm(long x, long y, long cx, long cy) => $"<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>";

    static string Ph(int id, string name, string phAttrs, string? xfrm, string? bodyPr, string? lst, string paragraph, bool locks = true) =>
        $"<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"{name}\"/><p:cNvSpPr>{(locks ? "<a:spLocks noGrp=\"1\"/>" : "")}</p:cNvSpPr><p:nvPr><p:ph {phAttrs}/></p:nvPr></p:nvSpPr><p:spPr>{(xfrm != null ? xfrm + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>" : "")}</p:spPr>" +
        $"<p:txBody>{bodyPr ?? "<a:bodyPr/>"}{lst ?? "<a:lstStyle/>"}{paragraph}</p:txBody></p:sp>";

    const string Prompt = "<a:p><a:r><a:rPr lang=\"it-IT\"/><a:t>Modifica il testo</a:t></a:r></a:p>";

    static string Footers(int firstId) =>
        Ph(firstId, "Data", "type=\"dt\" sz=\"half\" idx=\"10\"", null, null, null, "<a:p><a:endParaRPr lang=\"it-IT\"/></a:p>") +
        Ph(firstId + 1, "Piè di pagina", "type=\"ftr\" sz=\"quarter\" idx=\"11\"", null, null, null, "<a:p><a:endParaRPr lang=\"it-IT\"/></a:p>") +
        Ph(firstId + 2, "Numero diapositiva", "type=\"sldNum\" sz=\"quarter\" idx=\"12\"", null, null, null, "<a:p><a:fld id=\"{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}\" type=\"slidenum\"><a:rPr lang=\"it-IT\"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang=\"it-IT\"/></a:p>");

    static string Layout(string name, string type, string shapes) =>
        $"<p:sldLayout {Ns} type=\"{type}\" preserve=\"1\"><p:cSld name=\"{name}\"><p:spTree>{Grp}{shapes}{Footers(20)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>";

    static string Title(long? x = null, long y = 0, long cx = 0, long cy = 0, string? body = null, string? lst = null, string type = "title") =>
        Ph(2, "Titolo", $"type=\"{type}\"", x == null ? null : Xfrm(x.Value, y, cx, cy), body, lst, Prompt);

    static string Level(int n, int marL, int indent, int sz, bool bullet) =>
        $"<a:lvl{n}pPr marL=\"{marL}\" indent=\"{indent}\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:lnSpc><a:spcPct val=\"100000\"/></a:lnSpc><a:spcBef><a:spcPts val=\"600\"/></a:spcBef>" +
        (bullet ? "<a:buFont typeface=\"Arial\"/><a:buChar char=\"•\"/>" : "<a:buNone/>") +
        $"<a:defRPr sz=\"{sz}\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl{n}pPr>";

    static string Levels(bool bullet, int[] sizes) =>
        string.Concat(sizes.Select((s, i) => Level(i + 1, 342900 + i * 457200 - (bullet ? 0 : 342900), bullet ? -342900 : 0, s, bullet)));

    static string ThemeXml() => $"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Tema Prisma\"><a:themeElements>" +
        "<a:clrScheme name=\"Prisma\"><a:dk1><a:srgbClr val=\"1B1B1F\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"2B3A55\"/></a:dk2><a:lt2><a:srgbClr val=\"EEF1F5\"/></a:lt2>" +
        "<a:accent1><a:srgbClr val=\"FF7A45\"/></a:accent1><a:accent2><a:srgbClr val=\"4F8CFF\"/></a:accent2><a:accent3><a:srgbClr val=\"3ECF8E\"/></a:accent3><a:accent4><a:srgbClr val=\"F2B134\"/></a:accent4><a:accent5><a:srgbClr val=\"8E6CEF\"/></a:accent5><a:accent6><a:srgbClr val=\"E5484D\"/></a:accent6>" +
        "<a:hlink><a:srgbClr val=\"2F6FDE\"/></a:hlink><a:folHlink><a:srgbClr val=\"7A5BC0\"/></a:folHlink></a:clrScheme>" +
        "<a:fontScheme name=\"Prisma\"><a:majorFont><a:latin typeface=\"Segoe UI Semibold\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont><a:minorFont><a:latin typeface=\"Segoe UI\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>" +
        "<a:fmtScheme name=\"Prisma\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"85000\"/></a:schemeClr></a:solidFill></a:fillStyleLst>" +
        "<a:lnStyleLst><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"28575\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:lnStyleLst>" +
        "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"40000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>" +
        "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"90000\"/></a:schemeClr></a:solidFill></a:bgFillStyleLst></a:fmtScheme>" +
        "</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>";

    static string MasterXml(string layoutIds) => $"<p:sldMaster {Ns}><p:cSld><p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg><p:spTree>{Grp}" +
        Ph(2, "Titolo", "type=\"title\"", Xfrm(838200, 365125, 10515600, 1325563), "<a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"ctr\"><a:normAutofit/></a:bodyPr>", null, Prompt) +
        Ph(3, "Segnaposto testo", "type=\"body\" idx=\"1\"", Xfrm(838200, 1825625, 10515600, 4351338), "<a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"><a:normAutofit/></a:bodyPr>", null, Prompt) +
        Ph(4, "Data", "type=\"dt\" sz=\"half\" idx=\"2\"", Xfrm(838200, 6356350, 2743200, 365125), "<a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"ctr\"/>", "<a:lstStyle><a:lvl1pPr algn=\"l\"><a:defRPr sz=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"><a:tint val=\"75000\"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>", "<a:p><a:endParaRPr lang=\"it-IT\"/></a:p>") +
        Ph(5, "Piè di pagina", "type=\"ftr\" sz=\"quarter\" idx=\"3\"", Xfrm(4038600, 6356350, 4114800, 365125), "<a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"ctr\"/>", "<a:lstStyle><a:lvl1pPr algn=\"ctr\"><a:defRPr sz=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"><a:tint val=\"75000\"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>", "<a:p><a:endParaRPr lang=\"it-IT\"/></a:p>") +
        Ph(6, "Numero diapositiva", "type=\"sldNum\" sz=\"quarter\" idx=\"4\"", Xfrm(8610600, 6356350, 2743200, 365125), "<a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"ctr\"/>", "<a:lstStyle><a:lvl1pPr algn=\"r\"><a:defRPr sz=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"><a:tint val=\"75000\"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>", "<a:p><a:fld id=\"{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}\" type=\"slidenum\"><a:rPr lang=\"it-IT\"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang=\"it-IT\"/></a:p>") +
        $"</p:spTree></p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/><p:sldLayoutIdLst>{layoutIds}</p:sldLayoutIdLst>" +
        "<p:txStyles><p:titleStyle><a:lvl1pPr algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:lnSpc><a:spcPct val=\"95000\"/></a:lnSpc><a:spcBef><a:spcPct val=\"0\"/></a:spcBef><a:buNone/><a:defRPr sz=\"4000\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mj-lt\"/><a:ea typeface=\"+mj-ea\"/><a:cs typeface=\"+mj-cs\"/></a:defRPr></a:lvl1pPr></p:titleStyle>" +
        $"<p:bodyStyle>{Levels(true, new[] { 2800, 2400, 2000, 1800, 1800, 1800, 1800, 1800, 1800 })}</p:bodyStyle>" +
        $"<p:otherStyle><a:defPPr><a:defRPr lang=\"it-IT\"/></a:defPPr>{Levels(false, new[] { 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800 })}</p:otherStyle></p:txStyles></p:sldMaster>";

    const long L = 838200, W = 10515600;

    static readonly (string Name, string Type, string Shapes)[] Layouts =
    {
        ("Titolo", "title", Title(1524000, 1122363, 9144000, 2387600, "<a:bodyPr anchor=\"b\"/>", "<a:lstStyle><a:lvl1pPr algn=\"ctr\"><a:defRPr sz=\"5400\"/></a:lvl1pPr></a:lstStyle>", "ctrTitle") +
            Ph(3, "Sottotitolo", "type=\"subTitle\" idx=\"1\"", Xfrm(1524000, 3602038, 9144000, 1655762), null, "<a:lstStyle><a:lvl1pPr marL=\"0\" indent=\"0\" algn=\"ctr\"><a:buNone/><a:defRPr sz=\"2400\"/></a:lvl1pPr></a:lstStyle>", Prompt)),
        ("Titolo e contenuto", "obj", Title() + Ph(3, "Segnaposto contenuto", "idx=\"1\"", null, null, null, Prompt)),
        ("Intestazione di sezione", "secHead", Title(831850, 1709738, 10515600, 2852737, "<a:bodyPr anchor=\"b\"/>", "<a:lstStyle><a:lvl1pPr><a:defRPr sz=\"5400\"/></a:lvl1pPr></a:lstStyle>") +
            Ph(3, "Testo", "type=\"body\" idx=\"1\"", Xfrm(831850, 4589463, 10515600, 1500187), null, "<a:lstStyle><a:lvl1pPr marL=\"0\" indent=\"0\"><a:buNone/><a:defRPr sz=\"2400\"/></a:lvl1pPr></a:lstStyle>", Prompt)),
        ("Due contenuti", "twoObj", Title() + Ph(3, "Contenuto sinistro", "sz=\"half\" idx=\"1\"", Xfrm(L, 1825625, 5181600, 4351338), null, "<a:lstStyle><a:lvl1pPr><a:defRPr sz=\"2800\"/></a:lvl1pPr></a:lstStyle>", Prompt) +
            Ph(4, "Contenuto destro", "sz=\"half\" idx=\"2\"", Xfrm(6172200, 1825625, 5181600, 4351338), null, "<a:lstStyle><a:lvl1pPr><a:defRPr sz=\"2800\"/></a:lvl1pPr></a:lstStyle>", Prompt)),
        ("Solo titolo", "titleOnly", Title()),
        ("Vuota", "blank", ""),
    };

    static string Min(string root, string extra = "") => $"<p:{root} {Ns}{extra}/>";

    /// <summary>Costruisce il pacchetto del modello (master, layout, tema, una slide titolo vuota).</summary>
    public static byte[] Build()
    {
        using var ms = new MemoryStream();
        using (var doc = PresentationDocument.Create(ms, DocumentFormat.OpenXml.PresentationDocumentType.Presentation))
        {
            var pp = doc.AddPresentationPart();
            var master = pp.AddNewPart<SlideMasterPart>("rId1");
            var theme = master.AddNewPart<ThemePart>("rId100");
            theme.Theme = new A.Theme(ThemeXml());
            pp.AddPart(theme, "rId2");
            var ids = new StringBuilder();
            for (int i = 0; i < Layouts.Length; i++)
            {
                var lp = master.AddNewPart<SlideLayoutPart>($"rId{i + 1}");
                lp.SlideLayout = new P.SlideLayout(Layout(Layouts[i].Name, Layouts[i].Type, Layouts[i].Shapes));
                lp.AddPart(master, "rId1");
                ids.Append($"<p:sldLayoutId id=\"{2147483649L + i}\" r:id=\"rId{i + 1}\"/>");
            }
            // il rId della relazione al tema nel master deve essere diverso dai layout: lo ha creato con "rId100"
            master.SlideMaster = new P.SlideMaster(MasterXml(ids.ToString()));
            var props = pp.AddNewPart<PresentationPropertiesPart>("rId3"); props.PresentationProperties = new P.PresentationProperties(Min("presentationPr"));
            var view = pp.AddNewPart<ViewPropertiesPart>("rId4"); view.ViewProperties = new P.ViewProperties(Min("viewPr"));
            var ts = pp.AddNewPart<TableStylesPart>("rId5"); ts.TableStyleList = new A.TableStyleList("<a:tblStyleLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" def=\"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}\"/>");
            var lvls = string.Concat(Enumerable.Range(1, 9).Select(n => $"<a:lvl{n}pPr marL=\"{(n - 1) * 457200}\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1800\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl{n}pPr>"));
            pp.Presentation = new P.Presentation($"<p:presentation {Ns} saveSubsetFonts=\"1\"><p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst><p:sldSz cx=\"12192000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/><p:defaultTextStyle><a:defPPr><a:defRPr lang=\"it-IT\"/></a:defPPr>{lvls}</p:defaultTextStyle></p:presentation>");
            doc.PackageProperties.Creator = "PowerPresentations"; doc.PackageProperties.LastModifiedBy = "PowerPresentations"; doc.PackageProperties.Title = "Presentazione";
        }
        return ms.ToArray();
    }
}
