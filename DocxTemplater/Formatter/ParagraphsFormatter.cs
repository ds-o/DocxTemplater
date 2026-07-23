using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Formatter
{
    /// <summary>
    /// Renders a multi-line string value as real paragraphs: each line of the value becomes its own
    /// paragraph (a blank line an empty paragraph), inheriting the paragraph and run properties of the
    /// placeholder. Without this formatter, newlines in a value are rendered as soft line breaks
    /// (<c>w:br</c>) inside a single paragraph.
    /// </summary>
    internal class ParagraphsFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            if (type == typeof(string))
            {
                return prefix.Equals("paragraphs", StringComparison.OrdinalIgnoreCase) || prefix.Equals("para", StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {
            if (formatterContext.Value is not string value)
            {
                throw new OpenXmlTemplateException($"Formatter {formatterContext.Formatter} can only be applied to string objects - property {formatterContext.Placeholder}");
            }

            if (value.Length == 0)
            {
                // Keep the (possibly last) paragraph in place - removing it could leave a table cell empty.
                target.Text = string.Empty;
                return;
            }

            var lines = value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var hostParagraph = target.GetFirstAncestor<Paragraph>();

            // Single-line values need no splitting. Without a host paragraph, or inside an inline (run
            // level) content control, splitting would corrupt the surrounding structure - write the raw
            // value instead and let the caller's newline handling degrade to soft line breaks.
            if (lines.Length == 1 || hostParagraph == null || IsInsideInlineSdt(target, hostParagraph))
            {
                target.Text = value;
                target.Space = SpaceProcessingModeValues.Preserve;
                return;
            }

            var paragraphProperties = hostParagraph.GetFirstChild<ParagraphProperties>();
            var runProperties = target.GetFirstAncestor<Run>()?.RunProperties;

            // Text before the placeholder stays with the first line; text after it goes with the last.
            var split = hostParagraph.SplitAfterElement(target);
            var firstParagraph = split.OfType<Paragraph>().First();
            var lastParagraph = split.OfType<Paragraph>().Last();

            target.Text = lines[0];
            target.Space = SpaceProcessingModeValues.Preserve;

            // The split-off paragraph is a shallow clone without the host's paragraph properties; restore
            // them so the text after the placeholder keeps its original formatting.
            if (lastParagraph != firstParagraph && paragraphProperties != null && lastParagraph.GetFirstChild<ParagraphProperties>() == null)
            {
                lastParagraph.PrependChild(paragraphProperties.CloneNode(true));
            }

            var ownParagraphLines = lastParagraph != firstParagraph ? lines.Length - 1 : lines.Length;
            OpenXmlElement anchor = firstParagraph;
            for (int i = 1; i < ownParagraphLines; i++)
            {
                var paragraph = new Paragraph();
                if (paragraphProperties != null)
                {
                    paragraph.AppendChild(paragraphProperties.CloneNode(true));
                }
                if (lines[i].Length > 0)
                {
                    paragraph.AppendChild(CreateRun(runProperties, lines[i]));
                }
                anchor = anchor.InsertAfterSelf(paragraph);
            }

            if (lastParagraph != firstParagraph && lines[^1].Length > 0)
            {
                var run = CreateRun(runProperties, lines[^1]);
                var lastParagraphProperties = lastParagraph.GetFirstChild<ParagraphProperties>();
                if (lastParagraphProperties != null)
                {
                    lastParagraphProperties.InsertAfterSelf(run);
                }
                else
                {
                    lastParagraph.PrependChild(run);
                }
            }
        }

        private static Run CreateRun(RunProperties runProperties, string text)
        {
            var run = new Run();
            if (runProperties != null)
            {
                run.AppendChild((RunProperties)runProperties.CloneNode(true));
            }
            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            return run;
        }

        private static bool IsInsideInlineSdt(Text target, Paragraph hostParagraph)
        {
            for (var parent = target.Parent; parent != null && parent != hostParagraph; parent = parent.Parent)
            {
                if (parent is SdtElement)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
