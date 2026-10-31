using System.Buffers.Binary;
using Microsoft.Win32;

[assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")]
namespace PowerPresentations.Core.Rendering;

/// <summary>Metriche TrueType lette direttamente dal file del font (nessuna libreria di layout).</summary>
public sealed class FontFace
{
    public string Path { get; }
    /// <summary>Famiglia effettivamente trovata (può essere un ripiego).</summary>
    public string Family { get; private set; } = "";
    public double AscentEm { get; }
    public double DescentEm { get; }
    public double LineHeightEm { get; }
    readonly ushort[] _adv;
    readonly ushort _upem;
    readonly Dictionary<int, ushort> _cmap = new();
    ushort _defaultGlyph;

    FontFace(string path, byte[] d)
    {
        Path = path;
        var t = Tables(d);
        _upem = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(t["head"] + 18));
        var hh = t["hhea"];
        short asc = BinaryPrimitives.ReadInt16BigEndian(d.AsSpan(hh + 4));
        short desc = BinaryPrimitives.ReadInt16BigEndian(d.AsSpan(hh + 6));
        short gap = BinaryPrimitives.ReadInt16BigEndian(d.AsSpan(hh + 8));
        int nh = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(hh + 34));
        AscentEm = (double)asc / _upem;
        DescentEm = (double)-desc / _upem;
        LineHeightEm = (double)(asc - desc + gap) / _upem;
        _adv = new ushort[nh];
        for (int i = 0; i < nh; i++) _adv[i] = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(t["hmtx"] + 4 * i));
        ReadCmap(d, t["cmap"]);
        _defaultGlyph = 0;
    }

    static Dictionary<string, int> Tables(byte[] d)
    {
        int n = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(4));
        var r = new Dictionary<string, int>();
        for (int i = 0; i < n; i++)
        {
            int o = 12 + 16 * i;
            r[System.Text.Encoding.ASCII.GetString(d, o, 4)] = (int)BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(o + 8));
        }
        return r;
    }

    void ReadCmap(byte[] d, int cm)
    {
        int n = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(cm + 2));
        int best = -1;
        for (int i = 0; i < n; i++)
        {
            int pid = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(cm + 4 + 8 * i));
            int eid = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(cm + 6 + 8 * i));
            int off = (int)BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(cm + 8 + 8 * i));
            int fmt = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(cm + off));
            if (pid == 3 && (eid == 10 || eid == 1) && (fmt == 12 || fmt == 4)) { if (best < 0 || fmt == 12) best = cm + off; }
        }
        if (best < 0) return;
        int f = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(best));
        if (f == 4)
        {
            int seg = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(best + 6)) / 2;
            int endO = best + 14, startO = endO + seg * 2 + 2, deltaO = startO + seg * 2, rangeO = deltaO + seg * 2;
            for (int s = 0; s < seg; s++)
            {
                int end = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(endO + 2 * s));
                int start = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(startO + 2 * s));
                short delta = BinaryPrimitives.ReadInt16BigEndian(d.AsSpan(deltaO + 2 * s));
                int ro = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(rangeO + 2 * s));
                for (int c = start; c <= end && c != 0xFFFF; c++)
                {
                    int g;
                    if (ro == 0) g = (c + delta) & 0xFFFF;
                    else
                    {
                        int p = rangeO + 2 * s + ro + 2 * (c - start);
                        if (p + 2 > d.Length) continue;
                        g = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(p));
                        if (g != 0) g = (g + delta) & 0xFFFF;
                    }
                    if (g != 0) _cmap[c] = (ushort)g;
                }
            }
        }
        else
        {
            uint ng = BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(best + 12));
            for (uint g = 0; g < ng; g++)
            {
                int o = best + 16 + 12 * (int)g;
                uint s = BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(o)), e = BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(o + 4)), gl = BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(o + 8));
                if (e - s > 0x20000) continue;
                for (uint c = s; c <= e; c++) _cmap[(int)c] = (ushort)(gl + c - s);
            }
        }
    }

    public double Advance(int cp)
    {
        if (!_cmap.TryGetValue(cp, out var g)) g = _defaultGlyph;
        return (double)_adv[Math.Min(g, _adv.Length - 1)] / _upem;
    }

    public double Measure(string s, double size)
    {
        double w = 0;
        for (int i = 0; i < s.Length; i++)
        {
            int cp = char.ConvertToUtf32(s, i);
            if (char.IsHighSurrogate(s[i])) i++;
            w += Advance(cp);
        }
        return w * size;
    }

    // ---- registro dei font di sistema ----
    static Dictionary<string, string>? _reg;
    static readonly Dictionary<string, FontFace?> _cache = new(StringComparer.OrdinalIgnoreCase);
    static readonly object _lock = new();

    static Dictionary<string, string> Registry_()
    {
        if (_reg != null) return _reg;
        var r = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        void Read(RegistryKey? k, string dir)
        {
            if (k == null) return;
            foreach (var n in k.GetValueNames())
            {
                var v = k.GetValue(n) as string; if (v == null) continue;
                var name = n.Replace(" (TrueType)", "").Replace(" (OpenType)", "");
                if (!v.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase)) continue; // ponytail: niente .ttc/.otf CFF → fallback Arial
                r[name] = System.IO.Path.IsPathRooted(v) ? v : System.IO.Path.Combine(dir, v);
            }
        }
        var win = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        Read(Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"), win);
        Read(Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"), win);
        AddCloudFonts(r);
        return _reg = r;
    }

    // font cloud di Windows (es. Aptos): file con nome a hash in FontCache\CloudFonts\<famiglia>; il nome completo si legge dalla tabella 'name'
    static void AddCloudFonts(Dictionary<string, string> r)
    {
        var root = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "FontCache", "4", "CloudFonts");
        if (!Directory.Exists(root)) return;
        foreach (var f in Directory.EnumerateFiles(root, "*.ttf", SearchOption.AllDirectories))
        {
            try
            {
                // legge solo directory e tabella 'name' (i file possono pesare parecchio)
                using var fs = File.OpenRead(f);
                var hdr = new byte[12 + 16 * 40]; int got = fs.Read(hdr, 0, hdr.Length);
                var t = Tables(hdr); var nameOff = t["name"];
                int nt = BinaryPrimitives.ReadUInt16BigEndian(hdr.AsSpan(4)), nlen = 0;
                for (int k = 0; k < nt; k++) if (System.Text.Encoding.ASCII.GetString(hdr, 12 + 16 * k, 4) == "name") nlen = (int)BinaryPrimitives.ReadUInt32BigEndian(hdr.AsSpan(12 + 16 * k + 12));
                var d = new byte[nlen]; fs.Seek(nameOff, SeekOrigin.Begin); fs.ReadExactly(d, 0, nlen); int n = 0;
                int cnt = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(n + 2)), so = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(n + 4));
                for (int i = 0; i < cnt; i++)
                {
                    int o = n + 6 + 12 * i;
                    if (BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(o + 6)) != 4) continue;
                    int len = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(o + 8)), off = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(o + 10));
                    var name = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(o)) == 3 ? System.Text.Encoding.BigEndianUnicode.GetString(d, n + so + off, len) : System.Text.Encoding.ASCII.GetString(d, n + so + off, len);
                    r.TryAdd(name, f); break;
                }
            }
            catch (Exception e) when (e is IOException or KeyNotFoundException or ArgumentException or UnauthorizedAccessException) { /* font illeggibile: ignorato */ }
        }
    }

    /// <summary>Trova il file del font; ripiega su Arial. Lo stile mancante viene ignorato (ponytail: niente grassetto sintetico).</summary>
    public static FontFace Get(string? family, bool bold, bool italic)
    {
        family ??= "Arial";
        var key = $"{family}|{bold}|{italic}";
        lock (_lock)
        {
            if (_cache.TryGetValue(key, out var f) && f != null) return f;
            var reg = Registry_();
            var style = (bold ? " Bold" : "") + (italic ? " Italic" : "");
            string? path = null, found = null;
            // ponytail: font non installati (es. Aptos, font cloud) → Segoe UI, poi Arial; nessuna sostituzione tipografica fine
            foreach (var fam in new[] { family, "Segoe UI", "Arial" })
                foreach (var st in new[] { style, bold ? " Bold" : "", italic ? " Italic" : "", "" })
                    if (path == null && reg.TryGetValue(fam + st, out var p) && File.Exists(p)) { path = p; found = fam; }
            if (path == null) throw new FileNotFoundException("Nessun font TrueType di sistema trovato (nemmeno Arial).");
            f = new FontFace(path, File.ReadAllBytes(path)) { Family = found! };
            _cache[key] = f;
            return f;
        }
    }
}
