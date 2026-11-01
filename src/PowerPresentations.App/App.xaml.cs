using System.IO;
using System.Windows;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;


namespace PowerPresentations.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (_, ev) => { ThemedDialog.Notify(MainWindow, "Errore imprevisto", ev.Exception.Message + "\nIl file aperto non è stato modificato su disco.", DialogKind.Error); ev.Handled = true; };
        var a = e.Args;
        // modalità da riga di comando per i test visivi: --render-png <pptx> <cartella> [larghezza] | --export-pdf <pptx> <pdf>
        if (a.Length >= 3 && a[0] == "--render-png")
        {
            try
            {
                using var d = PptxDocument.Open(a[1]); Directory.CreateDirectory(a[2]);
                var r = new SlideRenderer(d); int w = a.Length > 3 ? int.Parse(a[3]) : 1280;
                for (int i = 0; i < d.Slides.Count; i++) SlideDrawing.SavePng(SlideDrawing.ToBitmap(r.Render(i), w), Path.Combine(a[2], $"slide{i + 1:00}.png"));
                Shutdown(0);
            }
            catch (Exception ex) { Console.Error.WriteLine(ex.Message); Shutdown(1); }
            return;
        }
        if (a.Length >= 3 && a[0] == "--export-pdf")
        {
            try { using var d = PptxDocument.Open(a[1]); PdfExporter.Export(d, a[2]); Shutdown(0); }
            catch (Exception ex) { Console.Error.WriteLine(ex.Message); Shutdown(1); }
            return;
        }
        if (a.Length >= 4 && a[0] == "--transition-frames")
        {
            // test della riproduzione: cattura la finestra di presentazione prima, a metà e dopo il passaggio alla slide a[2] (indice da 0)
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            var d = PptxDocument.Open(a[1]); int idx = int.Parse(a[2]);
            var c = new ShowController(d, new SlideRenderer(d), idx - 1);
            var w = new PresentationWindow(c) { WindowStyle = WindowStyle.SingleBorderWindow, ResizeMode = ResizeMode.NoResize, Width = 700, Height = 420, WindowStartupLocation = WindowStartupLocation.Manual, Left = 40, Top = 40 };
            w.Show();
            void At(int ms, Action act) { var t = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(ms) }; t.Tick += (_, _) => { t.Stop(); act(); }; t.Start(); }
            void Cap(string n) { var el = (FrameworkElement)w.Content; el.UpdateLayout(); var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap((int)el.ActualWidth, (int)el.ActualHeight, 96, 96, System.Windows.Media.PixelFormats.Pbgra32); rtb.Render(el); SlideDrawing.SavePng(rtb, a[3] + "-" + n + ".png"); }
            At(1200, () => { Cap("0-prima"); c.Next(); At(230, () => Cap("1-meta")); At(1800, () => { Cap("2-dopo"); d.Dispose(); Shutdown(0); }); });
            return;
        }
        ShutdownMode = ShutdownMode.OnMainWindowClose;   // chiudere la finestra principale chiude anche presentazione e vista relatore
        var win = new MainWindow(); 
        if (a.Length >= 1 && File.Exists(a[0])) win.OpenFile(a[0]);
        win.Show();
        if (a.Length >= 3 && a[1] == "--screenshot") win.ScheduleScreenshot(a[2], a.Length > 3 ? a[3] : null);
    }
}
