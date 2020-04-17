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
                ColorConsole.Write(".".DarkGray());
                if (string.IsNullOrEmpty(text))
                {
                    text = HttpUtility.UrlDecode(link).Split('/', StringSplitOptions.RemoveEmptyEntries).LastOrDefault().Replace("-", " ") + "?!";
                }

                results.Add(new Item($"{s}.{l}", text, link.Trim()));
            }
        }
    }
}
