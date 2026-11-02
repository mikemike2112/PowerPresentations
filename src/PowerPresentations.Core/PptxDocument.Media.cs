using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPresentations.Core;

/// <summary>Video e audio incorporati, come li scrive PowerPoint: p:pic con a:videoFile/a:audioFile + p14:media, immagine di anteprima (poster).</summary>
public sealed partial class PptxDocument
{
    /// <summary>ponytail: tetto a 50 MB perché annulla/ripristina tengono copie del documento in memoria → media collegati (file esterni).</summary>
    public const long MaxMediaBytes = 50L << 20;

    static readonly Dictionary<string, (string Kind, string Type)> MediaTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        [".mp4"] = ("video", "video/mp4"), [".m4v"] = ("video", "video/mp4"), [".mov"] = ("video", "video/quicktime"), [".wmv"] = ("video", "video/x-ms-wmv"),
        [".avi"] = ("video", "video/avi"), [".mpg"] = ("video", "video/mpeg"), [".mpeg"] = ("video", "video/mpeg"),
        [".mp3"] = ("audio", "audio/mpeg"), [".wav"] = ("audio", "audio/wav"), [".m4a"] = ("audio", "audio/mp4"), [".wma"] = ("audio", "audio/x-ms-wma")
    };

    public static string? MediaKind(string path) => MediaTypes.TryGetValue(Path.GetExtension(path), out var m) ? m.Kind : null;

    /// <summary>Inserisce un file video o audio con la sua immagine di anteprima. Ritorna l'id della forma.</summary>
    public int AddMedia(int slide, string path, byte[] posterPng, double x, double y, double w, double h)
    {
        if (!MediaTypes.TryGetValue(Path.GetExtension(path), out var mt)) throw new PptxFormatException("Formato multimediale non supportato (video: mp4, m4v, mov, wmv, avi, mpg; audio: mp3, wav, m4a, wma).");
        if (ImageType(posterPng) == null) throw new PptxFormatException("Anteprima non valida.");
        FileInfo fi;
        try { fi = new FileInfo(path); if (!fi.Exists) throw new FileNotFoundException(); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or NotSupportedException) { throw new PptxFormatException($"Impossibile leggere il file: {e.Message}", e); }
        if (fi.Length == 0) throw new PptxFormatException("Il file multimediale è vuoto.");
        if (fi.Length > MaxMediaBytes) throw new PptxFormatException($"Il file è troppo grande ({fi.Length >> 20} MB): il limite per l'incorporamento è {MaxMediaBytes >> 20} MB.");

        var sp = Slide(slide); int id = NextShapeId(sp);
        var ext = Path.GetExtension(path).ToLowerInvariant();
        var dp = _doc.CreateMediaDataPart(mt.Type, ext);
        try { using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read); dp.FeedData(fs); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { throw new PptxFormatException($"Impossibile leggere il file: {e.Message}", e); }
        string link = mt.Kind == "video" ? sp.AddVideoReferenceRelationship(dp).Id : sp.AddAudioReferenceRelationship(dp).Id;
        string media = sp.AddMediaReferenceRelationship(dp).Id;
        var ip = sp.AddImagePart(ImageType(posterPng)!);
        using (var ms = new MemoryStream(posterPng)) ip.FeedData(ms);
        string img = sp.GetIdOfPart(ip);
        string tag = mt.Kind == "video" ? "videoFile" : "audioFile", label = mt.Kind == "video" ? "Video" : "Audio";
        string descr = System.Security.SecurityElement.Escape(Path.GetFileName(path));
        var pic = new P.Picture($"<p:pic {NsDecl}><p:nvPicPr><p:cNvPr id=\"{id}\" name=\"{label} {id}\" descr=\"{descr}\"><a:hlinkClick r:id=\"\" action=\"ppaction://media\"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>" +
            $"<p:nvPr><a:{tag} r:link=\"{link}\"/><p:extLst><p:ext uri=\"{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}\"><p14:media xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" r:embed=\"{media}\"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>" +
            $"<p:blipFill><a:blip r:embed=\"{img}\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"{Emu(x)}\" y=\"{Emu(y)}\"/><a:ext cx=\"{Emu(w)}\" cy=\"{Emu(h)}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>");
        Tree(sp).AppendChild(pic);
        Dirty = true; return id;
    }
}
