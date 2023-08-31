using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Task = System.Threading.Tasks.Task;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Windows.Forms;

namespace AutoWordScreenshot {
    internal class Utils {

        private long MULTIPLIER = 9525;
        private long MAX_WIDTH = 6264000;

        /// <summary>
        /// Save the image file to a temporary location for pasting purposes.
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public string SaveTempImage(Image image) {
            try {
                string tempFilePath = Path.GetTempFileName();

                using (FileStream stream = new FileStream(tempFilePath, FileMode.Create)) {
                    image.Save(stream, ImageFormat.Png);
                }

                return tempFilePath;
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Fetch the word document, or trigger the creation if it does not exist.
        /// </summary>
        /// <param name="documentPath"></param>
        /// <returns></returns>
        public WordprocessingDocument GetDocument(string documentPath) {
            try {
                return PrepareDocument(WordprocessingDocument.Open(documentPath, true));
            } catch {
                return CreateDocument(documentPath);
            }
        }

        /// <summary>
        /// Create the word document with the appropriate pre-processing and preparation.
        /// </summary>
        /// <param name="documentPath"></param>
        /// <returns></returns>
        public WordprocessingDocument CreateDocument(string documentPath) {
            try {
                return PrepareDocument(WordprocessingDocument.Create(documentPath, WordprocessingDocumentType.Document, true));
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Save and insert the image into the prepared word document.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="entityId"></param>
        /// <param name="imageWidth"></param>
        /// <param name="imageHeight"></param>
        public void SaveImageToDocument(WordprocessingDocument document, string entityId, long imageWidth, long imageHeight) {
            long width = imageWidth * this.MULTIPLIER, height = imageHeight * this.MULTIPLIER;
            if (width > this.MAX_WIDTH) {
                var ratio = (height * 1.0m) / width;
                width = this.MAX_WIDTH;
                height = (long)(width * ratio);
            }

            var element = new Drawing(new DW.Inline(new DW.Extent() { Cx = width, Cy = height },
            new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties() { Id = (UInt32Value)11U, Name = "Picture 1" },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                new A.Graphic(new A.GraphicData(new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)10U, Name = "New Bitmap Image.jpg" },
                new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(
                new A.Blip(new A.BlipExtensionList(new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" })) { Embed = entityId, CompressionState = A.BlipCompressionValues.Print, },
                new A.Stretch(new A.FillRectangle())),
                new PIC.ShapeProperties(new A.Transform2D(
                new A.Offset() { X = 0L, Y = 0L },
                new A.Extents() { Cx = width, Cy = height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })) { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, EditId = "50D07946" });

            document.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        /// <summary>
        /// Get the default page margins for pre setup.
        /// </summary>
        /// <returns></returns>
        private PageMargin GetDefaultPageMargins() {
            return new PageMargin() { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 };
        }

        /// <summary>
        /// Prepare and process the word document with the relevant default setup.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private WordprocessingDocument PrepareDocument(WordprocessingDocument document) {
            if (document == null) {
                return null;
            }

            MainDocumentPart mainPart = document.MainDocumentPart;
            if (mainPart == null) {
                mainPart = document.AddMainDocumentPart();
            }

            if (mainPart.Document == null) {
                mainPart.Document = new Document();
            }

            if (mainPart.Document.Body == null) {
                Body body = new Body();
                SectionProperties sectionProperties = new SectionProperties();
                sectionProperties.Append(GetDefaultPageMargins());
                body.Append(sectionProperties);
                mainPart.Document.Body = body;
            }

            return document;
        }
    }
}
