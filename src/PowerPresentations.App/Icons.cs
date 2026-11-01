using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;

namespace PowerPresentations.App;

/// <summary>Set di icone proprio (geometrie vettoriali su griglia 24x24, tratto singolo): nessuna risorsa grafica di terzi.</summary>
public static class Icons
{
    static readonly Dictionary<string, string> D = new()
    {
        ["save"] = "M4,4 h13 l3,3 v13 h-16 z M8,4 v5 h8 v-5 M8,20 v-6 h8 v6",
        ["undo"] = "M8,6 L3,11 L8,16 M3,11 H15 A5,5 0 0 1 15,21 H11",
        ["redo"] = "M16,6 L21,11 L16,16 M21,11 H9 A5,5 0 0 0 9,21 H13",
        ["search"] = "M10,3 a7,7 0 1 1 -0.01,0 M15,15 L21,21",
        ["new"] = "M5,3 h9 l5,5 v13 h-14 z M14,3 v5 h5",
        ["open"] = "M3,7 h7 l2,2 h9 v10 h-18 z",
        ["pdf"] = "M5,3 h9 l5,5 v13 h-14 z M14,3 v5 h5 M8,15 h2 a1.5,1.5 0 0 0 0,-3 h-2 v6",
        ["slide-new"] = "M3,5 h14 v11 h-14 z M20,3 v7 M16.5,6.5 h7",
        ["duplicate"] = "M8,8 h12 v12 h-12 z M5,16 v-12 h12",
        ["trash"] = "M4,7 h16 M9,7 v-3 h6 v3 M6,7 l1,13 h10 l1,-13 M10,11 v6 M14,11 v6",
        ["hide"] = "M2,12 q10,-10 20,0 q-10,10 -20,0 M9,12 a3,3 0 1 0 6,0 a3,3 0 1 0 -6,0 M4,20 L20,4",
        ["play"] = "M7,4 L19,12 L7,20 z",
        ["play-here"] = "M3,3 h18 v13 h-18 z M10,6 L15,9.5 L10,13 z M12,16 v4 M8,20 h8",
        ["presenter"] = "M2,4 h12 v9 h-12 z M16,8 h6 v11 h-6 z M6,17 h4 M8,13 v4",
        ["table"] = "M3,4 h18 v16 h-18 z M3,9.5 h18 M3,15 h18 M9,4 v16 M15,4 v16",
        ["image"] = "M3,4 h18 v16 h-18 z M3,17 l5,-6 l4,4 l3,-3 l6,6 M16,8 a1.2,1.2 0 1 0 0.01,0",
        ["textbox"] = "M3,4 h18 v16 h-18 z M8,9 h8 M12,9 v8",
        ["shapes"] = "M3,13 h8 v8 h-8 z M17,3 a4.5,4.5 0 1 1 -0.01,0 M9,3 L13,10 H5 z",
        ["align-l"] = "M4,6 h16 M4,10 h10 M4,14 h16 M4,18 h10",
        ["align-c"] = "M4,6 h16 M7,10 h10 M4,14 h16 M7,18 h10",
        ["align-r"] = "M4,6 h16 M10,10 h10 M4,14 h16 M10,18 h10",
        ["align-j"] = "M4,6 h16 M4,10 h16 M4,14 h16 M4,18 h16",
        ["bullets"] = "M9,6 h11 M9,12 h11 M9,18 h11 M4,6 h.1 M4,12 h.1 M4,18 h.1",
        ["numbers"] = "M10,6 h10 M10,12 h10 M10,18 h10 M3,5 l1.5,-1 v4 M3,11 h3 l-3,4 h3",
        ["normal"] = "M3,4 h18 v16 h-18 z M9,4 v16",
        ["sorter"] = "M3,4 h8 v7 h-8 z M13,4 h8 v7 h-8 z M3,13 h8 v7 h-8 z M13,13 h8 v7 h-8 z",
        ["notes"] = "M3,3 h18 v12 h-18 z M3,19 h18 M3,22 h12",
        ["fit"] = "M4,9 v-5 h5 M15,4 h5 v5 M20,15 v5 h-5 M9,20 h-5 v-5",
        ["tr-none"] = "M4,6 h16 v12 h-16 z",
        ["tr-fade"] = "M4,6 h16 v12 h-16 z M8,9.5 h8 v5 h-8 z M4,12 h16",
        ["tr-push"] = "M3,12 h15 M12,6 l6,6 l-6,6 M21,5 v14",
        ["tr-wipe"] = "M4,6 h16 v12 h-16 z M12,6 v12 M15,9 l-3,3 l3,3",
        ["tr-cover"] = "M3,5 h13 v10 h-13 z M9,9 h12 v10 h-12 z",
        ["tr-pull"] = "M9,9 h12 v10 h-12 z M3,5 h13 v10 h-13 z",
        ["preview"] = "M3,4 h18 v13 h-18 z M10,7 L15,10.5 L10,14 z M9,21 h6",
        ["dir-l"] = "M19,12 H5 M11,6 L5,12 L11,18",
        ["dir-r"] = "M5,12 H19 M13,6 L19,12 L13,18",
        ["dir-u"] = "M12,19 V5 M6,11 L12,5 L18,11",
        ["dir-d"] = "M12,5 V19 M6,13 L12,19 L18,13",
        ["slow"] = "M4,18 a8,8 0 1 1 16,0 M12,18 L8,11", ["med"] = "M4,18 a8,8 0 1 1 16,0 M12,18 L12,9", ["fast"] = "M4,18 a8,8 0 1 1 16,0 M12,18 L17,11",
        ["paste"] = "M8,5 h8 v3 h-8 z M6,6 h-1 v15 h14 v-15 h-1 M9,13 h6 M9,17 h6",
        ["cut"] = "M6,3 L16,15 M18,3 L8,15 M5,17 a2.5,2.5 0 1 0 0.01,0 M19,17 a2.5,2.5 0 1 0 0.01,0",
        ["generic"] = "M4,4 h16 v16 h-16 z M8,12 h8",
        ["shield"] = "M12,3 L20,6 V12 C20,17 16.5,20 12,21 C7.5,20 4,17 4,12 V6 Z M8.5,12 L11,14.5 L15.5,9.5",
        ["more"] = "M5,12 h.1 M12,12 h.1 M19,12 h.1",
        ["moon"] = "M20,14 A8,8 0 1 1 10,4 A6,6 0 0 0 20,14 z",
        ["close"] = "M5,5 L19,19 M19,5 L5,19",
        ["min"] = "M5,12 h14",
        ["max"] = "M5,5 h14 v14 h-14 z",
        ["restore"] = "M8,8 h11 v11 h-11 z M5,15 v-10 h10",
        ["file-menu"] = "M3,6 h18 M3,12 h18 M3,18 h18",
        ["back"] = "M19,12 H5 M11,6 L5,12 L11,18",
    };

    /// <summary>Crea l'icona come Path (tratto col colore del tema, nessun riempimento).</summary>
    public static FrameworkElement Make(string name, double size = 22, Brush? stroke = null)
    {
        var p = new Path
        {
            Data = Geometry.Parse(D.TryGetValue(name, out var d) ? d : D["generic"]),
            StrokeThickness = 1.5, StrokeLineJoin = PenLineJoin.Round, StrokeStartLineCap = PenLineCap.Round, StrokeEndLineCap = PenLineCap.Round
        };
        if (stroke != null) p.Stroke = stroke; else p.SetResourceReference(Shape.StrokeProperty, "FgBrush");
        var c = new System.Windows.Controls.Canvas { Width = 24, Height = 24 }; c.Children.Add(p);
        return new System.Windows.Controls.Viewbox { Width = size, Height = size, Child = c };
    }
}
