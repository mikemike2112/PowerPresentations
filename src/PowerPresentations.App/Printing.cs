using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using PowerPresentations.Core;
using PowerPresentations.Core.Rendering;

namespace PowerPresentations.App;

/// <summary>Stampa: una diapositiva per pagina (esclude le nascoste, come PowerPoint), centrata e in scala.</summary>
public static class SlidePrint
{
    public static FixedDocument Build(PptxDocument doc, SlideRenderer r, double pageW, double pageH)
    {
        var fd = new FixedDocument();
        fd.DocumentPaginator.PageSize = new Size(pageW, pageH);
        for (int i = 0; i < doc.Slides.Count; i++)
        {
            if (doc.IsHidden(i)) continue;
            var dl = r.Render(i);
            double k = Math.Min((pageW - 48) / dl.Width, (pageH - 48) / dl.Height);
            var img = new Image { Source = SlideDrawing.ToBitmap(dl, (int)Math.Clamp(dl.Width * k * 2.5, 800, 3200)), Width = dl.Width * k, Height = dl.Height * k, Stretch = Stretch.Fill };
            var page = new FixedPage { Width = pageW, Height = pageH, Background = Brushes.White };
            FixedPage.SetLeft(img, (pageW - img.Width) / 2); FixedPage.SetTop(img, (pageH - img.Height) / 2);
            page.Children.Add(img);
            var pc = new PageContent(); ((System.Windows.Markup.IAddChild)pc).AddChild(page);
            fd.Pages.Add(pc);
        }
        return fd;
    }

    /// <summary>Scrive la stessa pagina che andrebbe in stampa su un file XPS (usato dal test, senza stampante né dialoghi).</summary>
    public static int WriteXps(PptxDocument doc, string path)
    {
        var fd = Build(doc, new SlideRenderer(doc), 1122.5, 793.7);   // A4 orizzontale a 96 dpi
        if (File.Exists(path)) File.Delete(path);
        using var pkg = System.IO.Packaging.Package.Open(path, FileMode.Create);
        using var xps = new System.Windows.Xps.Packaging.XpsDocument(pkg);
        System.Windows.Xps.Packaging.XpsDocument.CreateXpsDocumentWriter(xps).Write(fd.DocumentPaginator);
        return fd.Pages.Count;
    }
}
