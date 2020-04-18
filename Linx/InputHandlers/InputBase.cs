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
        private static readonly string[] Exclusions = new string[]
        {
        };

        public static IInput GetInstance(string file)
        {
            return Path.GetExtension(file).Equals(".pptx") ? (IInput)new DeckIn() : (IInput)new DocIn();
        }

        public virtual List<Item> ExtractLinks(object file)
        {
            throw new NotImplementedException();
        }

        protected void ParseLink(List<Item> results, int s, int l, string text, string link)
        {
            if (!string.IsNullOrWhiteSpace(link))
            {
                if (!Exclusions.Any(link.Contains))
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
