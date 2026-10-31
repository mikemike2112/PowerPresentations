using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

// uso: PdfRaster raster <pdf> <outDir> <larghezza>   |   PdfRaster compare <dirA> <dirB> <out.png> [larghezza]
if (args.Length >= 4 && args[0] == "raster") { await Raster(args[1], args[2], int.Parse(args[3])); return 0; }
if (args.Length >= 4 && args[0] == "compare") { Compare(args[1], args[2], args[3], args.Length > 4 ? int.Parse(args[4]) : 560); return 0; }
Console.Error.WriteLine("uso: raster <pdf> <outDir> <w> | compare <dirA> <dirB> <out.png> [w]"); return 1;

static async Task Raster(string pdf, string outDir, int width)
{
    Directory.CreateDirectory(outDir);
    var f = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(pdf));
    var doc = await PdfDocument.LoadFromFileAsync(f);
    for (uint i = 0; i < doc.PageCount; i++)
    {
        using var page = doc.GetPage(i);
        using var ms = new InMemoryRandomAccessStream();
        await page.RenderToStreamAsync(ms, new PdfPageRenderOptions { DestinationWidth = (uint)width, BackgroundColor = Windows.UI.Color.FromArgb(255, 255, 255, 255) });
        ms.Seek(0);
        using var fs = File.Create(Path.Combine(outDir, $"page{i + 1:00}.png"));
        await ms.AsStreamForRead().CopyToAsync(fs);
    }
}

static BitmapSource Load(string p)
{
    var b = new BitmapImage(); b.BeginInit(); b.CacheOption = BitmapCacheOption.OnLoad; b.UriSource = new Uri(Path.GetFullPath(p)); b.EndInit(); b.Freeze();
    return new FormatConvertedBitmap(b, PixelFormats.Bgra32, null, 0);
}

static void Compare(string da, string db, string output, int w)
{
    var a = Directory.GetFiles(da, "*.png").OrderBy(x => x).ToList(); var b = Directory.GetFiles(db, "*.png").OrderBy(x => x).ToList();
    int n = Math.Max(a.Count, b.Count); var rows = new List<(BitmapSource? A, BitmapSource? B, double Diff)>();
    for (int i = 0; i < n; i++)
    {
        var ia = i < a.Count ? Load(a[i]) : null; var ib = i < b.Count ? Load(b[i]) : null;
        rows.Add((ia, ib, ia != null && ib != null ? Diff(ia, ib) : -1));
    }
    int h = (int)(w * 9.0 / 16) + 4; int cols = 2;
    var dv = new DrawingVisual();
    using (var dc = dv.RenderOpen())
    {
        dc.DrawRectangle(Brushes.DimGray, null, new Rect(0, 0, cols * w + 3 * 6, n * (h + 18) + 6));
        for (int i = 0; i < n; i++)
        {
            double y = 6 + i * (h + 18);
            void Put(BitmapSource? s, double x)
            {
                if (s == null) return;
                double k = (double)w / s.PixelWidth; dc.DrawImage(s, new Rect(x, y + 14, w, s.PixelHeight * k));
            }
            Put(rows[i].A, 6); Put(rows[i].B, 12 + w);
            var ft = new FormattedText($"slide {i + 1}  [sinistra: PowerPoint | destra: clone]  differenza media pixel: {(rows[i].Diff < 0 ? "n/d" : rows[i].Diff.ToString("0.0") + "%")}", System.Globalization.CultureInfo.InvariantCulture, FlowDirection.LeftToRight, new Typeface("Segoe UI"), 11, Brushes.White, 1.0);
            dc.DrawText(ft, new Point(6, y));
        }
    }
    var rtb = new RenderTargetBitmap(cols * w + 18, n * (h + 18) + 6, 96, 96, PixelFormats.Pbgra32); rtb.Render(dv);
    var enc = new PngBitmapEncoder(); enc.Frames.Add(BitmapFrame.Create(rtb)); using var fs = File.Create(output); enc.Save(fs);
    Console.WriteLine(string.Join(" ", rows.Select((r, i) => $"s{i + 1}={(r.Diff < 0 ? "n/d" : r.Diff.ToString("0.0"))}%")));
}

static double Diff(BitmapSource a, BitmapSource b)
{
    int w = 160, h = 90;
    byte[] Px(BitmapSource s) { var t = new TransformedBitmap(s, new ScaleTransform((double)w / s.PixelWidth, (double)h / s.PixelHeight)); var buf = new byte[w * h * 4]; t.CopyPixels(buf, w * 4, 0); return buf; }
    var pa = Px(a); var pb = Px(b); long sum = 0;
    for (int i = 0; i < pa.Length; i++) if (i % 4 != 3) sum += Math.Abs(pa[i] - pb[i]);
    return sum * 100.0 / (pa.Length / 4 * 3 * 255);
}
