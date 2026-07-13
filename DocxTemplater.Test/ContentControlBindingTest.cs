using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    /// <summary>
    /// Binding of Word content controls (structured document tags) by a placeholder in their tag.
    /// A content control whose <c>w:tag</c> is a placeholder (e.g. <c>{{ds.Name}}</c>) has its content
    /// filled from the bound model, while the control itself and its tag are preserved so it can be
    /// filled again in a later pass.
    /// </summary>
    internal class ContentControlBindingTest
    {
        [Test]
        public void BlockContentControl_WithPlaceholderTag_IsFilledFromModel()
        {
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Name}}""/><w:showingPlcHdr/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>Click here to enter text.</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { Name = "World" }));

            var control = ContentControl(result, "{{ds.Name}}");
            Assert.Multiple(() =>
            {
                Assert.That(control.InnerText, Is.EqualTo("World"), "content is filled from the model");
                Assert.That(Tag(control), Is.EqualTo("{{ds.Name}}"), "the tag is preserved so the control stays addressable");
                Assert.That(control.SdtProperties.GetFirstChild<ShowingPlaceholder>(), Is.Null,
                    "the 'showing placeholder' flag is cleared so Word treats the value as real text");
            });
        }

        [Test]
        public void ContentControlTag_AppliesFormatter()
        {
            // The same formatter pipeline as text placeholders is reused, so any formatter works.
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Name}:ToUpper()}""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { Name = "world" }));

            Assert.That(ContentControl(result, "{{ds.Name}:ToUpper()}").InnerText, Is.EqualTo("WORLD"));
        }

        [Test]
        public void InlineRunContentControl_IsFilled()
        {
            const string body = @"
                <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:r><w:t xml:space=""preserve"">Reference: </w:t></w:r>
                  <w:sdt>
                    <w:sdtPr><w:tag w:val=""{{ds.Ref}}""/></w:sdtPr>
                    <w:sdtContent><w:r><w:t>placeholder</w:t></w:r></w:sdtContent>
                  </w:sdt>
                </w:p>";

            var result = Render(body, t => t.BindModel("ds", new { Ref = "R-42" }));

            Assert.Multiple(() =>
            {
                Assert.That(ContentControl(result, "{{ds.Ref}}").InnerText, Is.EqualTo("R-42"));
                Assert.That(result.InnerText, Is.EqualTo("Reference: R-42"));
            });
        }

        [Test]
        public void NonPlaceholderTag_IsLeftUnchanged()
        {
            // A content control with an ordinary (non-placeholder) tag must be ignored entirely,
            // so existing templates are unaffected.
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""JustATag""/><w:showingPlcHdr/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>original text</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { Name = "World" }));

            var control = ContentControl(result, "JustATag");
            Assert.Multiple(() =>
            {
                Assert.That(control.InnerText, Is.EqualTo("original text"), "content is untouched");
                Assert.That(control.SdtProperties.GetFirstChild<ShowingPlaceholder>(), Is.Not.Null,
                    "the placeholder flag is untouched");
            });
        }

        [Test]
        public void PlaceholderInsideContentControlContent_StillWorks_WithPlainTag()
        {
            // Regression: a placeholder placed in the *content* of a control with a plain tag keeps working
            // through the normal text pass (this is unchanged, pre-existing behaviour).
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""PlainTag""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>Hello {{ds.Name}}</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { Name = "World" }));

            Assert.That(ContentControl(result, "PlainTag").InnerText, Is.EqualTo("Hello World"));
        }

        [Test]
        public void UnboundTag_IsLeftUntouched_AndSecondPassFillsIt_WithoutClearingTheFirstPass()
        {
            // The deferred / multi-pass scenario: a document is generated (pass 1 fills what it can) and a
            // content control is filled later (pass 2), e.g. stamping a delivery date after review.
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Reference}}""/><w:showingPlcHdr/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>ref placeholder</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.DeliveryDate}}""/><w:showingPlcHdr/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>date placeholder</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            // Pass 1: only Reference is bound. DeliveryDate is unbound and must be left completely intact.
            var afterPass1 = RenderStream(CreateTemplate(body), Skip, t => t.BindModel("ds", new { Reference = "R-42" }));
            using (var doc1 = WordprocessingDocument.Open(CopyOf(afterPass1), false))
            {
                var b1 = doc1.MainDocumentPart.Document.Body;
                Assert.Multiple(() =>
                {
                    Assert.That(ContentControl(b1, "{{ds.Reference}}").InnerText, Is.EqualTo("R-42"), "Reference filled in pass 1");
                    var delivery = ContentControl(b1, "{{ds.DeliveryDate}}");
                    Assert.That(delivery.InnerText, Is.EqualTo("date placeholder"), "unbound control content is left untouched");
                    Assert.That(delivery.SdtProperties.GetFirstChild<ShowingPlaceholder>(), Is.Not.Null,
                        "unbound control keeps its placeholder flag for a later pass");
                });
            }

            // Pass 2: only DeliveryDate is bound. It must be filled WITHOUT clearing the value pass 1 wrote.
            var afterPass2 = RenderStream(afterPass1, Skip, t => t.BindModel("ds", new { DeliveryDate = "8 Jul 2026" }));
            using var doc2 = WordprocessingDocument.Open(afterPass2, false);
            var b2 = doc2.MainDocumentPart.Document.Body;
            Assert.Multiple(() =>
            {
                Assert.That(ContentControl(b2, "{{ds.DeliveryDate}}").InnerText, Is.EqualTo("8 Jul 2026"), "DeliveryDate filled in pass 2");
                Assert.That(ContentControl(b2, "{{ds.Reference}}").InnerText, Is.EqualTo("R-42"),
                    "the value written in pass 1 is NOT cleared by pass 2");
            });
        }

        [Test]
        public void ContentControlInLoop_IsFilledPerIteration()
        {
            const string body = @"
                <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:r><w:t>{{#ds.Items}}</w:t></w:r></w:p>
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Items}}""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>item</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>
                <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:r><w:t>{{/ds.Items}}</w:t></w:r></w:p>";

            var result = Render(body, t => t.BindModel("ds", new { Items = new[] { "a", "b", "c" } }));

            var controls = result.Descendants<SdtElement>().Where(c => Tag(c) == "{{ds.Items}}").ToList();
            string[] expected = ["a", "b", "c"];
            Assert.Multiple(() =>
            {
                Assert.That(controls, Has.Count.EqualTo(3), "one content control per loop item");
                Assert.That(controls.Select(c => c.InnerText).ToArray(), Is.EqualTo(expected));
            });
        }

        [Test]
        public void UnboundTag_ThrowsByDefault()
        {
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Missing}}""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            using var template = new DocxTemplate(CreateTemplate(body),
                new ProcessSettings { EnableContentControlTagBinding = true });
            template.BindModel("ds", new { Name = "World" });

            Assert.Throws<OpenXmlTemplateException>(() => template.Process());
        }

        [Test]
        public void PlaceholderTag_IsIgnored_WhenFlagDisabled()
        {
            // With the opt-in flag off (the default), a placeholder tag must be left completely alone -
            // no fill, no content change, no placeholder-flag removal.
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Name}}""/><w:showingPlcHdr/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>placeholder</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            MemoryStream result;
            using (var template = new DocxTemplate(CreateTemplate(body))) // default settings -> flag OFF
            {
                template.BindModel("ds", new { Name = "World" });
                var processed = template.Process();
                result = new MemoryStream();
                processed.Position = 0;
                processed.CopyTo(result);
                result.Position = 0;
            }

            using var doc = WordprocessingDocument.Open(result, false);
            var control = ContentControl(doc.MainDocumentPart.Document.Body, "{{ds.Name}}");
            Assert.Multiple(() =>
            {
                Assert.That(control.InnerText, Is.EqualTo("placeholder"), "the feature is inert when the flag is off");
                Assert.That(control.SdtProperties.GetFirstChild<ShowingPlaceholder>(), Is.Not.Null,
                    "the placeholder flag is left in place when the flag is off");
            });
        }

        [Test]
        public void ResolvedValue_WithNoTextTarget_IsSurfaced()
        {
            // An empty cell/row content control has nowhere to put a resolved value. Under ThrowException
            // that must surface as an error, not be silently dropped.
            const string body = @"
                <w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:tr>
                    <w:sdt>
                      <w:sdtPr><w:tag w:val=""{{ds.X}}""/></w:sdtPr>
                      <w:sdtContent><w:tc><w:tcPr><w:tcW w:w=""5000"" w:type=""dxa""/></w:tcPr><w:p/></w:tc></w:sdtContent>
                    </w:sdt>
                  </w:tr>
                </w:tbl>";

            using var template = new DocxTemplate(CreateTemplate(body),
                new ProcessSettings { EnableContentControlTagBinding = true });
            template.BindModel("ds", new { X = "value" });

            Assert.Throws<OpenXmlTemplateException>(() => template.Process());
        }

        [Test]
        public void HighlightErrorsMode_MarksUnboundContentControl()
        {
            const string body = @"
                <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:r><w:t>Intro</w:t></w:r></w:p>
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Missing}}""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var settings = new ProcessSettings { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument };
            var result = RenderStream(CreateTemplate(body), settings, t => t.BindModel("ds", new { Name = "World" }));

            using var doc = WordprocessingDocument.Open(result, false);
            var control = ContentControl(doc.MainDocumentPart.Document.Body, "{{ds.Missing}}");
            var shading = control.Descendants<Shading>().FirstOrDefault();
            Assert.Multiple(() =>
            {
                Assert.That(shading?.Fill?.Value, Is.EqualTo("FF0000"), "the unbound control content is highlighted as an error");
                Assert.That(control.InnerText, Is.EqualTo("x"), "the original content is highlighted in place, not cleared");
            });
        }

        [Test]
        public void MalformedPlaceholderTag_IsIgnored_AndDoesNotThrow()
        {
            // A tag that contains braces but is not a valid value binding (a malformed/block directive) must
            // be left alone, not abort the render - even in the default ThrowException mode.
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{}}""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>keep-a</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{/switch}}""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>keep-b</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { Name = "World" }));

            Assert.Multiple(() =>
            {
                Assert.That(ContentControl(result, "{{}}").InnerText, Is.EqualTo("keep-a"));
                Assert.That(ContentControl(result, "{{/switch}}").InnerText, Is.EqualTo("keep-b"));
            });
        }

        [Test]
        public void NullValue_LeavesControlUntouched()
        {
            // A bound-but-null value is treated like an unbound tag: the control is left unchanged (so a
            // shared model type carrying null members never wipes a control).
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Name}}""/><w:showingPlcHdr/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>placeholder</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { Name = (string)null }));

            var control = ContentControl(result, "{{ds.Name}}");
            Assert.Multiple(() =>
            {
                Assert.That(control.InnerText, Is.EqualTo("placeholder"), "null value leaves the content untouched");
                Assert.That(control.SdtProperties.GetFirstChild<ShowingPlaceholder>(), Is.Not.Null,
                    "null value leaves the placeholder flag in place");
            });
        }

        [Test]
        public void MultiPass_WithSharedModelType_DoesNotClearAnEarlierPass()
        {
            // Same guarantee as the anonymous-type deferred test, but with a shared DTO whose other member is
            // null (not absent) in each pass - which must NOT clear the value written by the previous pass.
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Reference}}""/><w:showingPlcHdr/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>ref</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.DeliveryDate}}""/><w:showingPlcHdr/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>date</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var pass1 = RenderStream(CreateTemplate(body), ProcessSettings.Default,
                t => t.BindModel("ds", new TwoFields { Reference = "R-42" }));
            var pass2 = RenderStream(pass1, ProcessSettings.Default,
                t => t.BindModel("ds", new TwoFields { DeliveryDate = "8 Jul 2026" }));

            using var doc = WordprocessingDocument.Open(pass2, false);
            var b = doc.MainDocumentPart.Document.Body;
            Assert.Multiple(() =>
            {
                Assert.That(ContentControl(b, "{{ds.DeliveryDate}}").InnerText, Is.EqualTo("8 Jul 2026"));
                Assert.That(ContentControl(b, "{{ds.Reference}}").InnerText, Is.EqualTo("R-42"),
                    "pass 1's value survives pass 2 even though ds.Reference is null in pass 2");
            });
        }

        [Test]
        public void MultiParagraphControl_CollapsesToSingleValue_NoStrayParagraph()
        {
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Name}}""/></w:sdtPr>
                  <w:sdtContent>
                    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>line1</w:t></w:r></w:p>
                    <w:p><w:r><w:t>line2</w:t></w:r></w:p>
                  </w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { Name = "World" }));

            var control = ContentControl(result, "{{ds.Name}}");
            Assert.Multiple(() =>
            {
                Assert.That(control.InnerText, Is.EqualTo("World"));
                Assert.That(control.Descendants<Paragraph>().Count(), Is.EqualTo(1), "no stray empty paragraph is left behind");
                Assert.That(control.Descendants<Bold>().Any(), Is.True, "the first run's formatting is preserved");
            });
        }

        [Test]
        public void NestedContentControl_OuterDoesNotCorruptInner()
        {
            // An outer bound control whose content also contains an inner (plain-tag) control must fill its
            // OWN text without touching the inner control's content.
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Outer}}""/></w:sdtPr>
                  <w:sdtContent>
                    <w:p>
                      <w:r><w:t xml:space=""preserve"">own </w:t></w:r>
                      <w:sdt>
                        <w:sdtPr><w:tag w:val=""InnerPlain""/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>KEEP</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
                  </w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { Outer = "OUTERVAL" }));

            Assert.Multiple(() =>
            {
                Assert.That(ContentControl(result, "InnerPlain").InnerText, Is.EqualTo("KEEP"),
                    "the inner control's content is not consumed by the outer fill");
                Assert.That(result.InnerText, Does.Contain("OUTERVAL"), "the outer control is still filled with its own value");
            });
        }

        [Test]
        public void ExpressionTag_IsEvaluated()
        {
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{(ds.A + ds.B)}}""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { A = 2, B = 3 }));

            Assert.That(ContentControl(result, "{{(ds.A + ds.B)}}").InnerText, Is.EqualTo("5"));
        }

        [Test]
        public void PlaceholderInTagAndContent_TagWins()
        {
            // Tag is a placeholder AND the content starts with a placeholder: the tag drives the fill and the
            // content placeholder is discarded without leaving a marker for the text pass to choke on.
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.A}}""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>{{ds.B}}</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { A = "AVAL", B = "BVAL" }));

            Assert.That(ContentControl(result, "{{ds.A}}").InnerText, Is.EqualTo("AVAL"));
        }

        [Test]
        public void EmptyContentControl_IsFilled()
        {
            const string body = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Name}}""/></w:sdtPr>
                  <w:sdtContent><w:p/></w:sdtContent>
                </w:sdt>";

            var result = Render(body, t => t.BindModel("ds", new { Name = "World" }));

            Assert.That(ContentControl(result, "{{ds.Name}}").InnerText, Is.EqualTo("World"));
        }

        [Test]
        public void ContentControlInHeader_IsFilled()
        {
            const string sdt = @"
                <w:sdt xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:sdtPr><w:tag w:val=""{{ds.Name}}""/></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>";

            var template = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(template, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph()));
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                headerPart.Header = new Header { InnerXml = sdt };
                wpDocument.Save();
            }
            template.Position = 0;

            var result = RenderStream(template, ProcessSettings.Default, t => t.BindModel("ds", new { Name = "World" }));

            using var doc = WordprocessingDocument.Open(result, false);
            var header = doc.MainDocumentPart.HeaderParts.First().Header;
            Assert.That(ContentControl(header, "{{ds.Name}}").InnerText, Is.EqualTo("World"));
        }

        // ---- helpers -------------------------------------------------------------------------------------

        // A model type shared across passes whose non-bound member is null (not absent) in each pass.
        private sealed class TwoFields
        {
            public string Reference { get; set; }
            public string DeliveryDate { get; set; }
        }

        private static readonly ProcessSettings Skip =
            new() { BindingErrorHandling = BindingErrorHandling.SkipBindingAndRemoveContent };

        private static MemoryStream CreateTemplate(string bodyInnerXml)
        {
            var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document { Body = new Body { InnerXml = bodyInnerXml } };
                wpDocument.Save();
            }
            memStream.Position = 0;
            return memStream;
        }

        private static Body Render(string bodyInnerXml, Action<DocxTemplate> bind)
        {
            var result = RenderStream(CreateTemplate(bodyInnerXml), ProcessSettings.Default, bind);
            using var doc = WordprocessingDocument.Open(result, false);
            // clone the body so it stays usable after the document is disposed
            return (Body)doc.MainDocumentPart.Document.Body.CloneNode(true);
        }

        private static MemoryStream RenderStream(Stream template, ProcessSettings settings, Action<DocxTemplate> bind)
        {
            // This fixture exercises the feature, so binding is enabled for all render helpers. The
            // flag-off behaviour is covered separately by PlaceholderTag_IsIgnored_WhenFlagDisabled.
            settings.EnableContentControlTagBinding = true;
            using var docTemplate = new DocxTemplate(template, settings);
            bind(docTemplate);
            var processed = docTemplate.Process();
            docTemplate.Validate();
            var copy = new MemoryStream();
            processed.Position = 0;
            processed.CopyTo(copy);
            copy.Position = 0;
            return copy;
        }

        private static MemoryStream CopyOf(MemoryStream source)
        {
            var copy = new MemoryStream(source.ToArray()) { Position = 0 };
            source.Position = 0;
            return copy;
        }

        private static SdtElement ContentControl(OpenXmlElement root, string tag)
        {
            return root.Descendants<SdtElement>().Single(c => Tag(c) == tag);
        }

        private static string Tag(SdtElement contentControl)
        {
            return contentControl.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
        }
    }
}
