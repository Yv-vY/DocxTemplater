using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxTemplater.Test
{
    [TestFixture]
    public class SubTemplateRawInsertTest
    {
        // The 'raw' argument inserts a sub-document verbatim, WITHOUT running the template engine over it.
        // This is for embedding user-authored / untrusted documents whose text may legitimately contain the
        // template syntax ({{ }}, {{#..}} loops) - e.g. code samples or template-injection findings in a
        // security report - which must survive as literal text rather than be interpreted (and removed).
        [Test]
        public void RawInsert_PreservesTemplateSyntaxVerbatim_AndImage()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");

            using var subDocStream = new MemoryStream();
            using (var subDocument = WordprocessingDocument.Create(subDocStream, WordprocessingDocumentType.Document))
            {
                var subMainPart = subDocument.AddMainDocumentPart();
                subMainPart.Document = new Document(new Body());
                var imagePart = subMainPart.AddImagePart(ImagePartType.Jpeg);
                using (var imageStream = new MemoryStream(imageBytes))
                {
                    imagePart.FeedData(imageStream);
                }
                var relationshipId = subMainPart.GetIdOfPart(imagePart);
                subMainPart.Document.Body.Append(
                    new Paragraph(new Run(new Text("Literal scalar {{ds.Name}} and loop {{#each users}}{{.name}}{{/each}} must survive.") { Space = SpaceProcessingModeValues.Preserve })),
                    new Paragraph(new Run(CreateInlineImage(relationshipId, 990000L, 792000L))));
                subMainPart.Document.Save();
            }

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("{{ds}:template('ds.SubDocument','raw')}")))));
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream, new ProcessSettings
            {
                // Even under the most aggressive error handling, raw content must not be touched.
                BindingErrorHandling = BindingErrorHandling.SkipBindingAndRemoveContent
            });
            // Note: "Name" and "users" are deliberately NOT bound - if the sub-document were processed they
            // would be treated as unbound and removed. Under 'raw' they must survive verbatim.
            docTemplate.BindModel("ds", new { SubDocument = subDocStream.ToArray() });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.Position = 0;

            using var document = WordprocessingDocument.Open(result, false);
            var mainDocumentPart = document.MainDocumentPart;
            var body = mainDocumentPart.Document.Body;

            // The template syntax inside the inserted document survived as literal text.
            Assert.That(body.InnerText, Does.Contain("{{ds.Name}}"));
            Assert.That(body.InnerText, Does.Contain("{{#each users}}"));
            Assert.That(body.InnerText, Does.Contain("{{/each}}"));

            // The embedded image is still preserved and resolvable in the raw insert.
            var blip = body.Descendants<A.Blip>().SingleOrDefault();
            Assert.That(blip, Is.Not.Null, "the inserted image should be present in the raw-merged document");
            Assert.That(blip.Embed?.Value, Is.Not.Null.And.Not.Empty);
            Assert.That(mainDocumentPart.GetPartById(blip.Embed.Value), Is.InstanceOf<ImagePart>());
        }

        private static Drawing CreateInlineImage(string relationshipId, long cx, long cy)
        {
            var picture = new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = "image.jpg" },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())),
                new PIC.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = cx, Cy = cy }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

            var inline = new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.DocProperties { Id = 1U, Name = "Picture 1" },
                new A.Graphic(new A.GraphicData(picture) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            };
            return new Drawing(inline);
        }
    }
}
