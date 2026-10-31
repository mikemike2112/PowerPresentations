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
        DispatcherUnhandledException += (_, ev) => { MessageBox.Show("Errore imprevisto: " + ev.Exception.Message + "\nIl file aperto non è stato modificato su disco.", "PowerPresentations", MessageBoxButton.OK, MessageBoxImage.Error); ev.Handled = true; };
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
        var win = new MainWindow(); 
        if (a.Length >= 1 && File.Exists(a[0])) win.OpenFile(a[0]);
        win.Show();
        if (a.Length >= 3 && a[1] == "--screenshot") win.ScheduleScreenshot(a[2], a.Length > 3 ? a[3] : null);
    }
}
