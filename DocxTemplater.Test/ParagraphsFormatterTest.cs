using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    /// <summary>
    /// The 'paragraphs' (alias 'para') formatter renders each line of a multi-line value as its own
    /// real paragraph (a blank line as an empty paragraph) instead of the default soft line breaks
    /// (&lt;w:br/&gt;) inside a single paragraph. Soft breaks keep all lines in one paragraph, which
    /// breaks paragraph-based styling, numbering and downstream tooling that works per paragraph.
    /// </summary>
    internal class ParagraphsFormatterTest
    {
        [Test]
        public void MultiLineValue_CreatesOneParagraphPerLine()
        {
            using var processed = Render(new Paragraph(new Run(new Text("{{ds.Notes}:paragraphs}"))), "Line1\nLine2\nLine3");
            var body = processed.MainDocumentPart.Document.Body;

            Assert.That(body.Elements<Paragraph>().Select(p => p.InnerText), Is.EqualTo(["Line1", "Line2", "Line3"]));
            Assert.That(body.Descendants<Break>(), Is.Empty, "lines must become paragraphs, not soft breaks");
        }

        [Test]
        public void BlankLineAndCrLf_BecomeEmptyParagraph()
        {
            using var processed = Render(new Paragraph(new Run(new Text("{{ds.Notes}:paragraphs}"))), "Alinea1\r\n\r\nAlinea2");
            var body = processed.MainDocumentPart.Document.Body;

            Assert.That(body.Elements<Paragraph>().Select(p => p.InnerText), Is.EqualTo(["Alinea1", "", "Alinea2"]));
            Assert.That(body.Descendants<Break>(), Is.Empty);
        }

        [Test]
        public void TrailingNewline_YieldsTrailingEmptyParagraph()
        {
            using var processed = Render(new Paragraph(new Run(new Text("{{ds.Notes}:paragraphs}"))), "Last\n");
            var body = processed.MainDocumentPart.Document.Body;

            Assert.That(body.Elements<Paragraph>().Select(p => p.InnerText), Is.EqualTo(["Last", ""]));
        }

        [Test]
        public void SurroundingText_StaysOnFirstAndLastLine()
        {
            using var processed = Render(new Paragraph(new Run(new Text("Start {{ds.Notes}:para} End"))), "a\nb");
            var body = processed.MainDocumentPart.Document.Body;

            Assert.That(body.Elements<Paragraph>().Select(p => p.InnerText), Is.EqualTo(["Start a", "b End"]));
        }

        [Test]
        public void SingleLineValue_StaysInline()
        {
            using var processed = Render(new Paragraph(new Run(new Text("Start {{ds.Notes}:paragraphs} End"))), "plain");
            var body = processed.MainDocumentPart.Document.Body;

            Assert.That(body.Elements<Paragraph>().Select(p => p.InnerText), Is.EqualTo(["Start plain End"]));
        }

        [Test]
        public void HostParagraphAndRunFormatting_CarryOverToAllLines()
        {
            var template = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Quote" }),
                new Run(new RunProperties(new Bold()), new Text("{{ds.Notes}:paragraphs}")));

            using var processed = Render(template, "a\nb\nc");
            var paragraphs = processed.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

            Assert.That(paragraphs, Has.Count.EqualTo(3));
            Assert.That(paragraphs.Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value),
                Is.All.EqualTo("Quote"), "every line must keep the host paragraph style");
            Assert.That(paragraphs.SelectMany(p => p.Descendants<Run>()).Select(r => r.RunProperties?.Bold),
                Is.All.Not.Null, "every line must keep the placeholder run formatting");
        }

        [Test]
        public void EmptyValue_InTableCell_KeepsCellParagraph()
        {
            const string content = @"
<w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:tblPr><w:tblW w:w=""0"" w:type=""auto""/></w:tblPr>
  <w:tblGrid><w:gridCol w:w=""5000""/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w=""5000"" w:type=""dxa""/></w:tcPr>
      <w:p><w:r><w:t>{{ds.Notes}:paragraphs}</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>";

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document { Body = new Body { InnerXml = content } };
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Notes = string.Empty });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var processed = WordprocessingDocument.Open(result, false);
            var cell = processed.MainDocumentPart.Document.Body.Descendants<TableCell>().Single();
            Assert.That(cell.Elements<Paragraph>().Count(), Is.EqualTo(1), "the cell must keep its paragraph");
        }

        [Test]
        public void InsideLoop_ExpandsPerIteration()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("{{#ds.Items}}"))),
                    new Paragraph(new Run(new Text("{{.Value}:paragraphs}"))),
                    new Paragraph(new Run(new Text("{{/ds.Items}}")))));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Items = new[] { new { Value = "a\nb" }, new { Value = "c\nd" } } });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var processed = WordprocessingDocument.Open(result, false);
            var texts = processed.MainDocumentPart.Document.Body.Elements<Paragraph>()
                .Select(p => p.InnerText)
                .Where(t => t.Length > 0)
                .ToList();
            Assert.That(texts, Is.EqualTo(["a", "b", "c", "d"]));
        }

        [Test]
        public void InsideInlineContentControl_FallsBackToSoftBreaks()
        {
            const string content = @"
<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:sdt>
    <w:sdtPr><w:tag w:val=""notes""/></w:sdtPr>
    <w:sdtContent><w:r><w:t>{{ds.Notes}:paragraphs}</w:t></w:r></w:sdtContent>
  </w:sdt>
</w:p>";

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document { Body = new Body { InnerXml = content } };
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Notes = "a\nb" });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var processed = WordprocessingDocument.Open(result, false);
            var body = processed.MainDocumentPart.Document.Body;

            // Splitting a run-level content control would duplicate it; the value degrades to soft breaks.
            Assert.That(body.Elements<Paragraph>().Count(), Is.EqualTo(1));
            Assert.That(body.Descendants<Break>().Count(), Is.EqualTo(1));
            Assert.That(body.InnerText, Does.Contain("a").And.Contain("b"));
        }

        private static WordprocessingDocument Render(Paragraph templateParagraph, string notes)
        {
            var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(templateParagraph));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Notes = notes });
            var result = docTemplate.Process();
            docTemplate.Validate();
            return WordprocessingDocument.Open(result, false);
        }
    }
}
