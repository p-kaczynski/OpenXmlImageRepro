using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXmlImageRepro
{
    class Program
    {
        static void Main()
        {
            const string pathToDocx = "test.docx";
            const string pathToPng = "test.png";

            using (var doc = WordprocessingDocument.Create(pathToDocx, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();

                // Upload image
                var imagePart = main.AddImagePart(ImagePartType.Png);

                using (var stream = File.OpenRead(pathToPng))
                    imagePart.FeedData(stream);


                var imageId = main.GetIdOfPart(imagePart);

                // Create empty document with single paragraph and run
                main.Document = new Document();

                var body = main.Document.AppendChild(new Body());

                var para = body.AppendChild(new Paragraph());

                var run = para.AppendChild(new Run());

                var image = new Drawing(
                 new DW.Inline(
                     new DW.Extent { Cx = 2257425L, Cy = 2143125L },
                     new DW.EffectExtent
                     {
                         LeftEdge = 0L,
                         TopEdge = 0L,
                         RightEdge = 9525L,
                         BottomEdge = 9525L
                     },
                     new DW.DocProperties
                     {
                         Id = (UInt32Value)1U,
                         Name = "Picture 1"
                     },
                     new DW.NonVisualGraphicFrameDrawingProperties(
                         new A.GraphicFrameLocks { NoChangeAspect = true }),
                     new A.Graphic(
                         new A.GraphicData(
                             new PIC.Picture(
                                 new PIC.NonVisualPictureProperties(
                                     new PIC.NonVisualDrawingProperties
                                     {
                                         Id = (UInt32Value)0U,
                                         Name = "test.png"
                                     },
                                     new PIC.NonVisualPictureDrawingProperties()),
                                 new PIC.BlipFill(
                                     new A.Blip(
                                         new A.BlipExtensionList(
                                             new A.BlipExtension(
                                                 new UseLocalDpi
                                                 {
                                                     Val = false
                                                 })
                                             {
                                                 Uri =
                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                             })
                                     )
                                     {
                                         Embed = imageId,
                                         //CompressionState =
                                         //A.BlipCompressionValues.Print
                                     },
                                     new A.Stretch(
                                         new A.FillRectangle())),
                                 new PIC.ShapeProperties(
                                     new A.Transform2D(
                                         new A.Offset { X = 0L, Y = 0L },
                                         new A.Extents { Cx = 2257425L, Cy = 2143125L }),
                                     new A.PresetGeometry(
                                         new A.AdjustValueList()
                                     )
                                     { Preset = A.ShapeTypeValues.Rectangle }))
                         )
                         { Uri = "https://schemas.openxmlformats.org/drawingml/2006/picture" })
                 )
                 {
                     DistanceFromTop = (UInt32Value)0U,
                     DistanceFromBottom = (UInt32Value)0U,
                     DistanceFromLeft = (UInt32Value)0U,
                     DistanceFromRight = (UInt32Value)0U,
                     //EditId = "50D07946",
                     //AnchorId = "5E7DDB1C"
                 });

                run.AppendChild(image);
            }

        }
    }
}
