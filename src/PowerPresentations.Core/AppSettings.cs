using System.Text.Json;

namespace PowerPresentations.Core;

/// <summary>Impostazioni dell'app (%APPDATA%\PowerPresentations\settings.json). Un file illeggibile o corrotto non blocca l'avvio: si usano i predefiniti.</summary>
public sealed class AppSettings
{
    public bool Dark { get; set; } = true;

    /// <summary>Percorso predefinito; la variabile PP_SETTINGS lo sostituisce (test e screenshot non toccano le impostazioni reali).</summary>
    public static string DefaultPath => Environment.GetEnvironmentVariable("PP_SETTINGS") is { Length: > 0 } o ? o : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PowerPresentations", "settings.json");

    public static AppSettings Load(string? path = null)
    {
        path ??= DefaultPath;
        try { return JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(path)) ?? new AppSettings(); }
        catch (Exception e) when (e is IOException or JsonException or UnauthorizedAccessException or NotSupportedException) { return new AppSettings(); }
    }

    public void Save(string? path = null)
    {
        path ??= DefaultPath;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            var tmp = path + ".tmp"; File.WriteAllText(tmp, JsonSerializer.Serialize(this)); File.Move(tmp, path, true);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* impostazioni non salvate: l'app continua */ }
    }
}
