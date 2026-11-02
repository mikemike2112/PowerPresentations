using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using PowerPresentations.Core;

namespace PowerPresentations.App;

/// <summary>Inserimento di video e audio.</summary>
public partial class MainWindow
{
    /// <summary>Immagine di anteprima disegnata da noi (nessun fotogramma estratto dal file → ponytail: poster generico, niente frame reale).</summary>
    static byte[] MakePoster(string kind)
    {
        bool video = kind == "video"; int w = video ? 640 : 192, h = video ? 360 : 192;
        var g = new Grid { Width = w, Height = h, Background = new SolidColorBrush(Color.FromRgb(0x26, 0x26, 0x2b)) };
        g.Children.Add(Icons.Make(video ? "play" : "audio", video ? 120 : 110, Brushes.White));
        g.Measure(new Size(w, h)); g.Arrange(new Rect(0, 0, w, h)); g.UpdateLayout();
        var rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32); rtb.Render(g);
        var enc = new PngBitmapEncoder(); enc.Frames.Add(BitmapFrame.Create(rtb)); using var ms = new MemoryStream(); enc.Save(ms); return ms.ToArray();
    }

    void OnInsertMedia(string kind)
    {
        if (_doc == null) return;
        var dlg = new OpenFileDialog
        {
            Title = kind == "video" ? "Inserisci video" : "Inserisci audio",
            Filter = kind == "video" ? "Video (*.mp4;*.m4v;*.mov;*.wmv;*.avi;*.mpg;*.mpeg)|*.mp4;*.m4v;*.mov;*.wmv;*.avi;*.mpg;*.mpeg" : "Audio (*.mp3;*.wav;*.m4a;*.wma)|*.mp3;*.wav;*.m4a;*.wma"
        };
        if (dlg.ShowDialog(this) != true) return;
        InsertMediaFile(dlg.FileName);
    }

    public int InsertMediaFile(string path)
    {
        if (_doc == null) return 0;
        var kind = PptxDocument.MediaKind(path);
        if (kind == null) { ThemedDialog.Notify(this, "File non inseribile", "Formato multimediale non supportato.", DialogKind.Warning); return 0; }
        var (sw, sh) = _doc.SlideSize;
        double w = kind == "video" ? Math.Min(480, sw * 0.6) : 48, h = kind == "video" ? w * 9 / 16 : 48;
        int id = 0;
        if (Try(() => id = _doc.AddMedia(_cur, path, MakePoster(kind), (sw - w) / 2, (sh - h) / 2, w, h), "File multimediale non inseribile")) { RefreshCurrent(); View.SelectedId = id; }
        return id;
    }
}
