namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web;

    using ColoredConsole;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    internal class DeckEx
    {
        internal static List<Item> ExtractDeckLinks(string file)
        {
            var app = new Application();
            var presentation = app.Presentations.Open(file, MsoTriState.msoCTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

            var results = ParseSlides(presentation);

            presentation.Close();
            presentation.NAR();
            app.Quit();
            app.NAR();

            ColorConsole.WriteLine(Environment.NewLine);
            return results;
        }

        private static List<Item> ParseSlides(Presentation presentation)
        {
            var results = new List<Item>();
            var s = 1;
            foreach (Slide slide in presentation.Slides)
            {
                ColorConsole.Write(".".Green());
                var links = slide.Hyperlinks;
                if (links.Count > 0)
                {
                    ParseLinks(results, s, links);
                }

                links.NAR();
                slide.NAR();
                s++;
            }

            return results;
        }

        private static void ParseLinks(List<Item> results, int s, Hyperlinks links)
        {
            var l = 1;
            foreach (Hyperlink link in links)
            {
                ParseLink(results, s, l, link);
                l++;
            }

            results.Add(new Item(string.Empty, string.Empty, string.Empty));
        }

        private static void ParseLink(List<Item> results, int s, int l, Hyperlink link)
        {
            ColorConsole.Write(".".DarkGray());
            var text = link.TextToDisplay?.Trim();
            if (string.IsNullOrEmpty(text))
            {
                text = HttpUtility.UrlDecode(link.Address).Split('/', StringSplitOptions.RemoveEmptyEntries).LastOrDefault().Replace("-", " ") + "?!";
            }

            results.Add(new Item($"{s}.{l}", text, link.Address.Trim()));
            link.NAR();
        }
    }
}
