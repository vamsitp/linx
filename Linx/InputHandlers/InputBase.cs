namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Web;

    using ColoredConsole;

    public interface IInput
    {
        List<Item> ExtractLinks(object file);
    }

    public abstract class InputBase : IInput
    {
        private static readonly Dictionary<string, IInput> Inputs = new Dictionary<string, IInput>
        {
            { ".pptx", (IInput)new DeckIn() },
            { ".docx", (IInput)new DeckIn() }
        };

        public static List<string> Exclusions { get; set; }

        public static IInput GetInstance(string file)
        {
            return Inputs[Path.GetExtension(file)];
        }

        public virtual List<Item> ExtractLinks(object file)
        {
            throw new NotImplementedException();
        }

        protected void ParseLinks(List<Item> results, int s, List<Item> items)
        {
            if (items?.Count > 0)
            {
                var l = 1;
                foreach (Item item in items)
                {
                    ParseLink(results, s, l, item.Text, item.Link);
                    l++;
                }

                results.Add(new Item(string.Empty, string.Empty, string.Empty));
            }
        }

        protected void ParseLink(List<Item> results, int s, int l, string text, string link)
        {
            if (!string.IsNullOrWhiteSpace(link))
            {
                var sanitizedLink = link?.SanitizeLink();
                if (!(Exclusions?.Any(e => sanitizedLink.Contains(e, StringComparison.OrdinalIgnoreCase)) == true))
                {
                    var sanitizedText = text?.SanitizeText(sanitizedLink);
                    if (!results.Any(r => r.Text.Equals(sanitizedText, StringComparison.Ordinal) && r.Link.Equals(sanitizedLink, StringComparison.Ordinal)))
                    {
                        ColorConsole.Write(l.ToString());
                        results.Add(new Item($"{s}.{l}", sanitizedText, sanitizedLink));
                    }
                    else
                    {
                        ColorConsole.Write(l.ToString().DarkGray());
                    }
                }
            }
        }
    }
}
