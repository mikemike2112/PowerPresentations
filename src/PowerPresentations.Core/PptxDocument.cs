using System.IO.Packaging;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PowerPresentations.Core;

/// <summary>File non valido, corrotto o non supportato: il messaggio è pensato per l'utente.</summary>
public sealed class PptxFormatException : Exception
{
    public PptxFormatException(string message, Exception? inner = null) : base(message, inner) { }
}

/// <summary>
/// Documento .pptx aperto. Il modello è una VISTA sul package Open XML originale (caricato in memoria):
/// il salvataggio riscrive solo le part toccate, tutto il resto (part, relazioni, elementi sconosciuti) resta invariato.
/// </summary>
public sealed partial class PptxDocument : IDisposable
{
    MemoryStream _ms;
    PresentationDocument _doc;
    public string? FilePath { get; private set; }
    public bool Dirty { get; private set; }
    public PresentationDocument Package => _doc;
    public PresentationPart PresentationPart => _doc.PresentationPart!;

    PptxDocument(MemoryStream ms, PresentationDocument doc, string? path) { _ms = ms; _doc = doc; FilePath = path; }

    public static PptxDocument Open(string path)
    {
        byte[] bytes;
        try { bytes = File.ReadAllBytes(path); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { throw new PptxFormatException($"Impossibile leggere il file: {e.Message}", e); }
        var d = FromBytes(bytes);
        d.FilePath = path;
        return d;
    }

    public static PptxDocument FromBytes(byte[] bytes)
    {
        var ms = new MemoryStream();
        ms.Write(bytes); ms.Position = 0;
        try
        {
            var doc = PresentationDocument.Open(ms, true, new OpenSettings { AutoSave = false });
            var pp = doc.PresentationPart ?? throw new PptxFormatException("Il file non contiene una presentazione (manca ppt/presentation.xml).");
            if (pp.Presentation == null) throw new PptxFormatException("presentation.xml non valido.");
            return new PptxDocument(ms, doc, null);
        }
        catch (PptxFormatException) { ms.Dispose(); throw; }
        catch (Exception e) when (e is DocumentFormat.OpenXml.Packaging.OpenXmlPackageException or InvalidDataException or FileFormatException or System.Xml.XmlException or IOException or NullReferenceException or InvalidOperationException or ArgumentException)
        {
            ms.Dispose();
            throw new PptxFormatException("Il file non è una presentazione .pptx valida o è danneggiato: " + e.Message, e);
        }
    }

    /// <summary>Nuova presentazione dal modello incorporato (tema, master e layout).</summary>
    public static PptxDocument New()
    {
        var d = FromBytes(TemplateBuilder.Build());
        d.AddSlide(d.Layouts[0], -1);
        d.Dirty = true;
        return d;
    }

    public void MarkDirty() => Dirty = true;

    /// <summary>Copia completa dello stato corrente (per annulla/ripristina).</summary>
    public byte[] Snapshot()
    {
        _doc.Save();
        using var o = new MemoryStream();
        using (var c = _doc.Clone(o)) { }
        return o.ToArray();
    }

    /// <summary>Ripristina uno stato precedente ottenuto con <see cref="Snapshot"/>; i riferimenti a part ottenuti prima non sono più validi.</summary>
    public void Restore(byte[] state)
    {
        var ms = new MemoryStream(); ms.Write(state); ms.Position = 0;
        var doc = PresentationDocument.Open(ms, true, new OpenSettings { AutoSave = false });
        _doc.Dispose(); _ms.Dispose(); _doc = doc; _ms = ms; Dirty = true;
    }

    // ---- slide ----
    public SlideIdList? SlideIds => PresentationPart.Presentation!.SlideIdList;

    public IReadOnlyList<SlidePart> Slides =>
        (SlideIds?.Elements<SlideId>() ?? Enumerable.Empty<SlideId>())
            .Select(i => i.RelationshipId?.Value is { } r && PresentationPart.TryGetPartById(r, out var p) ? p as SlidePart : null)
            .Where(p => p != null).Select(p => p!).ToList();

    /// <summary>Dimensioni slide in punti.</summary>
    public (double W, double H) SlideSize
    {
        get
        {
            var s = PresentationPart.Presentation!.SlideSize;
            double cx = s?.Cx?.Value ?? 9144000, cy = s?.Cy?.Value ?? 6858000;
            return (cx / 12700.0, cy / 12700.0);
        }
    }

    /// <summary>Legge il contenuto XML di una part senza forzare il caricamento del DOM dell'SDK (le part non toccate restano intatte).</summary>
    public static XElement Xml(OpenXmlPart p)
    {
        if (p.RootElement != null) return XElement.Parse(p.RootElement.OuterXml);
        using var s = p.GetStream(FileMode.Open, FileAccess.Read);
        return XElement.Load(s);
    }

    // ---- salvataggio ----
    public void Save() { if (FilePath == null) throw new InvalidOperationException("Nessun percorso: usare SaveAs."); SaveAs(FilePath); }

    /// <summary>Salvataggio atomico: scrive su file temporaneo nella stessa cartella, poi sostituisce. Il file sorgente non si corrompe se fallisce.</summary>
    public void SaveAs(string path)
    {
        var full = Path.GetFullPath(path);
        var tmp = full + ".tmp" + Environment.ProcessId;
        try
        {
            _doc.Save();
            using (var c = _doc.Clone(tmp, false)) { }
            if (File.Exists(full)) File.Replace(tmp, full, null); else File.Move(tmp, full);
        }
        catch
        {
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch (IOException) { }
            throw;
        }
        FilePath = full; Dirty = false;
    }

    public void Dispose() { _doc.Dispose(); _ms.Dispose(); }
}
