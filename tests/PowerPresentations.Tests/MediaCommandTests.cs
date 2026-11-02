using DocumentFormat.OpenXml.Packaging;
using PowerPresentations.Core;

namespace PowerPresentations.Tests;

public class MediaCommandTests
{
    /// <summary>WAV PCM 8 bit mono di 0,2 s: file audio vero, non un segnaposto.</summary>
    public static byte[] Wav()
    {
        var ms = new MemoryStream(); var w = new BinaryWriter(ms); int n = 1600;
        w.Write("RIFF"u8.ToArray()); w.Write(36 + n); w.Write("WAVEfmt "u8.ToArray()); w.Write(16); w.Write((short)1); w.Write((short)1); w.Write(8000); w.Write(8000); w.Write((short)1); w.Write((short)8);
        w.Write("data"u8.ToArray()); w.Write(n); for (int i = 0; i < n; i++) w.Write((byte)(128 + 40 * Math.Sin(i / 6.0)));
        return ms.ToArray();
    }

    public static string WriteFile(string name, byte[] data) { var p = Path.Combine(Env.Work, name); File.WriteAllBytes(p, data); return p; }
    public static byte[] FakeMp4() => new byte[] { 0, 0, 0, 0x18, (byte)'f', (byte)'t', (byte)'y', (byte)'p', (byte)'i', (byte)'s', (byte)'o', (byte)'m', 0, 0, 2, 0, (byte)'i', (byte)'s', (byte)'o', (byte)'m', (byte)'m', (byte)'p', (byte)'4', (byte)'1', 0, 0, 0, 8, (byte)'f', (byte)'r', (byte)'e', (byte)'e' };

    [Fact]
    public void AddMedia_AudioAndVideo_ValidRelationships_PosterAndBytesKept()
    {
        using var d = PptxDocument.New();
        var wav = Wav(); var a = WriteFile("clip.wav", wav); var v = WriteFile("clip.mp4", FakeMp4());
        int ia = d.AddMedia(0, a, ComEditFixture.Png, 100, 100, 48, 48), iv = d.AddMedia(0, v, ComEditFixture.Png, 200, 100, 320, 180);
        Assert.NotEqual(ia, iv); Assert.Equal("audio", PptxDocument.MediaKind(a)); Assert.Equal("video", PptxDocument.MediaKind(v));
        using var r = CommandTests.Reload(d, "media");
        var sl = r.Slides[0];
        Assert.Equal(2, sl.DataPartReferenceRelationships.Count(x => x is MediaReferenceRelationship));
        Assert.Single(sl.DataPartReferenceRelationships, x => x is AudioReferenceRelationship); Assert.Single(sl.DataPartReferenceRelationships, x => x is VideoReferenceRelationship);
        var audio = sl.DataPartReferenceRelationships.First(x => x is AudioReferenceRelationship).DataPart;
        using var s = audio.GetStream(); var ms = new MemoryStream(); s.CopyTo(ms); Assert.Equal(wav, ms.ToArray());
        Assert.Equal(2, sl.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Count(p => p.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.InnerXml.Contains("p14:media")));
    }

    [Fact]
    public void AddMedia_RejectsBadInput_WithoutChangingTheDocument()
    {
        using var d = PptxDocument.New(); var before = d.Snapshot().Length;
        var txt = WriteFile("x.txt", new byte[] { 1 }); var big = Path.Combine(Env.Work, "big.mp4");
        using (var f = File.Create(big)) f.SetLength(PptxDocument.MaxMediaBytes + 1);
        var empty = WriteFile("empty.mp3", Array.Empty<byte>());
        Assert.Throws<PptxFormatException>(() => d.AddMedia(0, txt, ComEditFixture.Png, 0, 0, 10, 10));
        Assert.Throws<PptxFormatException>(() => d.AddMedia(0, big, ComEditFixture.Png, 0, 0, 10, 10));
        Assert.Throws<PptxFormatException>(() => d.AddMedia(0, empty, ComEditFixture.Png, 0, 0, 10, 10));
        Assert.Throws<PptxFormatException>(() => d.AddMedia(0, Path.Combine(Env.Work, "manca.mp4"), ComEditFixture.Png, 0, 0, 10, 10));
        Assert.Throws<PptxFormatException>(() => d.AddMedia(0, WriteFile("ok.wav", Wav()), new byte[] { 1, 2, 3 }, 0, 0, 10, 10));
        Assert.Empty(d.Slides[0].DataPartReferenceRelationships); Assert.Equal(before, d.Snapshot().Length);
    }

    [Fact]
    public void CopyPaste_Media_CopiesTheMediaBytes_AcrossDocuments()
    {
        using var d = PptxDocument.New(); var wav = Wav();
        int id = d.AddMedia(0, WriteFile("c.wav", wav), ComEditFixture.Png, 100, 100, 48, 48);
        var clip = ShapeClip.FromBytes(d.CopyShapes(0, new[] { id }).ToBytes())!;
        using var dst = PptxDocument.New(); dst.PasteShapes(0, clip);
        using var r = CommandTests.Reload(dst, "media-paste");
        var audio = r.Slides[0].DataPartReferenceRelationships.First(x => x is AudioReferenceRelationship).DataPart;
        using var s = audio.GetStream(); var ms = new MemoryStream(); s.CopyTo(ms); Assert.Equal(wav, ms.ToArray());
        Assert.Equal(2, r.Slides[0].DataPartReferenceRelationships.Count(x => x is MediaReferenceRelationship || x is AudioReferenceRelationship));
        Assert.Single(r.Slides[0].DataPartReferenceRelationships.Select(x => x.DataPart.Uri).Distinct());   // una sola part dati condivisa
    }
}
