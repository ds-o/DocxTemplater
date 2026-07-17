using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace DocxTemplater.Test
{
    internal class SubTemplateFormatterTest
    {

        [Test]
        public void SubTemplateTest()
        {
            var template = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                              <w:pPr>
                                <w:pBdr>
                                  <w:bottom w:val=""double"" w:sz=""6"" w:space=""1"" w:color=""auto""/>
                                </w:pBdr>
                              </w:pPr>
                              <w:r>
                                <w:t>Test {{ds.Name}} {{ds.Number}}</w:t>
                              </w:r>
                            </w:p>";

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("Start of Document")),
                    new Break(),
                    new Run(new Text("{{#ds.Items}}"))
                ),
            new Paragraph(
                    new Run(new Text("{{.Name}}")),
                    new Run(new Text("{{.}:T('ds.Template')}"))
                ),
            new Paragraph(
                new Run(new Text("{{/ds.Items}}"))
            )
            ));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds",
                new
                {
                    Template = template,
                    Items = new[]
                        {
                            new {Name = "Item1 ", Number = 55 },
                            new {Name = "Item2 ", Number = 96 }
                        }
                });
            var result = docTemplate.Process();
            //docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            //check values have been replaced
            Assert.That(body.InnerText, Is.EqualTo("Start of DocumentItem1 Test Item1  55Item2 Test Item2  96"));

            //check paragraphs have been added
            Assert.That(body.ChildElements.OfType<Paragraph>().Count(), Is.EqualTo(5));
            Assert.That(body.ChildElements.Any(e => e is not Paragraph), Is.False);
        }

        [Test]
        public void SubTemplateInlineTestParagraph()
        {
            var template = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                              <w:pPr>
                                <w:pBdr>
                                  <w:bottom w:val=""double"" w:sz=""6"" w:space=""1"" w:color=""auto""/>
                                </w:pBdr>
                              </w:pPr>
                              <w:bookmarkStart w:id=""1"" w:name=""ImportantSection""/>
                              <w:r>
                                <w:t>First run</w:t>
                              </w:r>
                              <w:bookmarkEnd w:id=""1""/>
                              <w:r>
                                <w:t>Second run {{Name}} {{Number}}</w:t>
                              </w:r>
                            </w:p>";

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("Start of Document"), new Break()),
                    new Run(new Text("{{ds.Item1}:T('ds.Template')}")),
                    new Run(new Text("Pre-tag2 {{ds.Item2}:T('ds.Template')} post-tag2"), new Break()),
                    new Run(new Text("End of Document"))
                )
            ));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(
                memStream,
                new ProcessSettings
                {
                    MergeSubTemplatesParagraph = true
                }
            );
            docTemplate.BindModel("ds",
                new
                {
                    Template = template,
                    Item1 = new { Name = "Item1 ", Number = 55 },
                    Item2 = new { Name = "Item2 ", Number = 33 }
                });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            //check values have been replaced
            Assert.That(body.InnerText, Is.EqualTo("Start of DocumentFirst runSecond run Item1  55Pre-tag2 First runSecond run Item2  33 post-tag2End of Document"));

            //check paragraphs have been merged
            Assert.That(body.ChildElements.OfType<Paragraph>().Count(), Is.EqualTo(1));
            Assert.That(body.ChildElements.Any(e => e is not Paragraph), Is.False);

            var para = body.GetFirstChild<Paragraph>()!;
            Assert.That(para.ChildElements.Count, Is.EqualTo(12));
            Assert.That(para.ChildElements.OfType<Run>().Count(), Is.EqualTo(8));
            Assert.That(para.ChildElements.OfType<BookmarkStart>().Count(), Is.EqualTo(2));
            Assert.That(para.ChildElements.OfType<BookmarkEnd>().Count(), Is.EqualTo(2));

            // Any Paragraph properties should have been removed
            Assert.That(para.ChildElements.OfType<ParagraphProperties>().Any(), Is.False);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void SubTemplateInlineTestBody(bool multiParagraph)
        {
            //create the sub-template
            using var memStreamSub = new MemoryStream();
            using var wpSubTemplate = WordprocessingDocument.Create(memStreamSub, WordprocessingDocumentType.Document);
            MainDocumentPart mainPartSub = wpSubTemplate.AddMainDocumentPart();
            mainPartSub.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds.value1}}")))));

            if (multiParagraph)
            {
                mainPartSub.Document.Body.AppendChild(new Paragraph(new Run(new Text("{{ds.value2}}"))));
            }

            wpSubTemplate.Save();
            memStreamSub.Position = 0;

            //create the template
            using var memStream = new MemoryStream();
            using var wpTemplate = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpTemplate.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Before {{.}:T('SubTemplate')} after")))));
            wpTemplate.Save();
            memStream.Position = 0;

            //render it
            using var docTemplate = new DocxTemplate(memStream,
                new ProcessSettings {
                    MergeSubTemplatesParagraph = true
                });
            docTemplate.BindModel("ds",
                new
                {
                    Value1 = "cat",
                    Value2 = "dog",
                    SubTemplate = memStreamSub
                });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Console.WriteLine(body.InnerText);
            // ...

            if (multiParagraph)
            {
                //template paragraph containing the sub-template tag is broken in multiple paragraphs
                Assert.That(body.ChildElements.OfType<Paragraph>().Count(), Is.EqualTo(4));
            }
            else
            {
                Assert.That(body.ChildElements.OfType<Paragraph>().Count(), Is.EqualTo(1));
            }
        }

        [TestCase(true)]
        [TestCase(false)]
        public void SubTemplateInsertDocxDocument(bool bindAsStream)
        {
            // create the document to insert - it can itself contain placeholders
            using var subDocStream = new MemoryStream();
            using (var subDocument = WordprocessingDocument.Create(subDocStream, WordprocessingDocumentType.Document))
            {
                var subMainPart = subDocument.AddMainDocumentPart();
                subMainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("First inserted paragraph for {{ds.Name}}"))),
                    new Paragraph(new Run(new Text("Second inserted paragraph")))));
            }

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Start of Document"))),
                new Paragraph(new Run(new Text("{{ds}:template('ds.SubDocument')}"))),
                new Paragraph(new Run(new Text("End of Document")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new
            {
                Name = "John",
                SubDocument = bindAsStream ? new MemoryStream(subDocStream.ToArray()) : (object)subDocStream.ToArray()
            });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            // the body content of the sub document is inserted at the placeholder position
            // and placeholders in the sub document are resolved against the bound model
            Assert.That(body.InnerText, Is.EqualTo("Start of DocumentFirst inserted paragraph for JohnSecond inserted paragraphEnd of Document"));
        }

    }
}
