using System.Globalization;

namespace DocxTemplater
{
    public class ProcessSettings
    {

        /// <summary>
        /// Output culture of the document
        /// </summary>
        public CultureInfo Culture { get; set; } = CultureInfo.CurrentUICulture;

        public BindingErrorHandling BindingErrorHandling { get; set; } = BindingErrorHandling.ThrowException;

        /// <summary>
        /// When enabled, this option removes leading or trailing newlines around template directives (e.g., {{#...}}, {{/}})
        /// from the final output. This allows templates to be more readable without affecting rendered formatting.
        /// default: false
        /// </summary>
        public bool IgnoreLineBreaksAroundTags { get; set; }

        /// <summary>
        /// When enabled, content controls whose tag is a placeholder (e.g. {{ds.Name}})
        /// are filled from the model. Default: false.
        /// </summary>
        public bool EnableContentControlTagBinding { get; set; }

        /// <summary>
        /// When enabled, the content of a sub-template made of a single top-level paragraph will be added to the
        /// destination paragraph instead of being inserted as a whole new paragraph. This allows to control the format
        /// of the sub document fragment via the target paragraph format (text alignment, etc.) and avoid to have an
        /// additional line return in the rendered document.
        /// This is especially useful for inline, short, templates
        /// default: false
        /// </summary>
        public bool InlineSubTemplates { get; set; }

        public static ProcessSettings Default => new();
    }
}
