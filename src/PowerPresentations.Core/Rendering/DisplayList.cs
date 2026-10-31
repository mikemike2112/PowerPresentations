namespace PowerPresentations.Core.Rendering;

/// <summary>Colore RGBA a 8 bit per canale.</summary>
public readonly record struct Rgba(byte R, byte G, byte B, byte A = 255)
{
    public string Hex => $"{R:X2}{G:X2}{B:X2}";
    public static Rgba FromHex(string h) => new(Convert.ToByte(h[..2], 16), Convert.ToByte(h[2..4], 16), Convert.ToByte(h[4..6], 16));
    public static readonly Rgba White = new(255, 255, 255), Black = new(0, 0, 0);
}

public readonly record struct RectD(double X, double Y, double W, double H);

public enum SegKind { Move, Line, Cubic, Close }

/// <summary>Segmento di tracciato in punti (1 pt = 1/72"), già con rotazione e ribaltamento applicati.</summary>
public readonly record struct Seg(SegKind Kind, double X1 = 0, double Y1 = 0, double X2 = 0, double Y2 = 0, double X3 = 0, double Y3 = 0);

public sealed record GradStop(double Pos, Rgba Color);
public abstract record Fill;
public sealed record SolidFill(Rgba Color) : Fill;
public sealed record GradientFill(IReadOnlyList<GradStop> Stops, double AngleDeg) : Fill;
public sealed record Stroke(Rgba Color, double Width, double[]? Dash = null);

/// <summary>Elemento della display list. Indipendente da WPF: schermo, miniature, presentazione e PDF la consumano uguale.</summary>
public abstract record DrawItem { public int ShapeId { get; init; } }
public sealed record PathItem(IReadOnlyList<Seg> Segs, Fill? Fill, Stroke? Stroke) : DrawItem;
public sealed record TextItem(string Text, double X, double Baseline, string Font, double Size, bool Bold, bool Italic, Rgba Color, double Rot = 0, double Cx = 0, double Cy = 0) : DrawItem;
public sealed record ImageItem(byte[]? Data, string ContentType, RectD Bounds, RectD Crop, double Rot, bool FlipH, bool FlipV) : DrawItem;

public sealed class DisplayList
{
    public double Width { get; init; }
    public double Height { get; init; }
    public Fill? Background { get; set; }
    public List<DrawItem> Items { get; } = new();
    /// <summary>Forme ignorate perché malformate (il file non viene modificato).</summary>
    public List<string> Warnings { get; } = new();
    /// <summary>Rimpicciolimento calcolato per le forme con normAutofit (solo con RenderOptions.RecomputeAutofit): id → (fontScale, lnSpcReduction).</summary>
    public Dictionary<int, (double Scale, double Reduction)> Fits { get; } = new();
    /// <summary>Riquadri delle forme di primo livello della slide (per selezione e modifica): id → rettangolo.</summary>
    public List<(int Id, RectD Box, bool HasText)> Hits { get; } = new();
}
