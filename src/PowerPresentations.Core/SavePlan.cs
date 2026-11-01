namespace PowerPresentations.Core;

public enum SaveChoice { Save, DontSave, Cancel }

/// <summary>Logica (senza UI) della richiesta di salvataggio: validazione del nome, percorso risultante, sovrascrittura.</summary>
public static class SavePlan
{
    public const string Extension = ".pptx";
    static readonly string[] Reserved = { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };

    /// <summary>Ritorna il messaggio d'errore per un nome file non valido, o null se va bene.</summary>
    public static string? ValidateName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "Scrivi un nome per il file.";
        name = name.Trim();
        if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return "Il nome non può contenere \\ / : * ? \" < > | né caratteri di controllo.";
        if (name.EndsWith('.') || name.EndsWith(' ')) return "Il nome non può terminare con un punto o uno spazio.";
        var stem = Path.GetFileNameWithoutExtension(name);
        if (Reserved.Contains(stem, StringComparer.OrdinalIgnoreCase)) return $"«{stem}» è un nome riservato di Windows.";
        if (name.Length + Extension.Length > 200) return "Il nome è troppo lungo.";
        return null;
    }

    /// <summary>Errore sulla cartella (vuota, relativa, inesistente), o null.</summary>
    public static string? ValidateFolder(string? folder)
    {
        if (string.IsNullOrWhiteSpace(folder)) return "Scegli una cartella.";
        try { if (!Path.IsPathFullyQualified(folder)) return "Il percorso della cartella deve essere completo."; }
        catch (ArgumentException) { return "Il percorso della cartella non è valido."; }
        return Directory.Exists(folder) ? null : "La cartella non esiste.";
    }

    /// <summary>Percorso completo: aggiunge .pptx se manca.</summary>
    public static string Combine(string folder, string name)
    {
        name = name.Trim();
        if (!name.EndsWith(Extension, StringComparison.OrdinalIgnoreCase)) name += Extension;
        return Path.Combine(folder, name);
    }

    /// <summary>Esito della verifica: errore da mostrare, oppure se serve chiedere conferma di sovrascrittura.</summary>
    public static (string? Error, bool NeedsOverwriteConfirm, string Path) Check(string? folder, string? name)
    {
        var e = ValidateName(name) ?? ValidateFolder(folder);
        if (e != null) return (e, false, "");
        var p = Combine(folder!, name!);
        return (null, File.Exists(p), p);
    }

    /// <summary>Solo Salva scrive; Non salvare e Annulla non scrivono mai.</summary>
    public static bool Writes(SaveChoice c) => c == SaveChoice.Save;
    /// <summary>Annulla interrompe l'operazione in corso (chiusura, Nuovo, Apri); Non salvare la prosegue.</summary>
    public static bool Proceeds(SaveChoice c) => c != SaveChoice.Cancel;
}
