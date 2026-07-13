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

        public static ProcessSettings Default => new();
    }
}
