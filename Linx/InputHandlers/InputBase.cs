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

        protected void ParseLink(List<Item> results, int s, int l, string text, string link)
        {
            if (!string.IsNullOrWhiteSpace(link))
            {
                if (!(Exclusions?.Any(link.Contains) == true))
                {
                    if (string.IsNullOrEmpty(text))
                    {
                        text = HttpUtility.UrlDecode(link).Split('/', StringSplitOptions.RemoveEmptyEntries).LastOrDefault().Replace("-", " ") + "?!";
                    }

                    if (!results.Any(x => x.Text.Equals(text, StringComparison.Ordinal) && x.Link.Equals(link, StringComparison.Ordinal)))
                    {
                        ColorConsole.Write(l.ToString().DarkGray());
                        results.Add(new Item($"{s}.{l}", text, link));
                    }
                    else
                    {
                        ColorConsole.Write(l.ToString().DarkRed());
                    }
                }
            }
        }
    }
}
