namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web;
    using ColoredConsole;
    using Microsoft.Office.Interop.Word;

    internal class DocEx
    {
        internal static List<Item> ExtractDocLinks(object file)
        {
            var results = new List<Item>();
            var app = new Application();
            object missing = Type.Missing;
            var doc = app.Documents.Open(
                            ref file,
                            ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing);

            ParsePages(results, app, doc);

            doc.Close();
            doc.NAR();
            app.Quit();
            app.NAR();

            ColorConsole.WriteLine(Environment.NewLine);
            return results;
        }

        private static void ParsePages(List<Item> results, Application app, Document doc)
        {
            var p = 1;
            var pageCount = doc.ComputeStatistics(WdStatistic.wdStatisticPages);
            var lastPageEnd = 0;
            for (long i = 0; i < pageCount; i++)
            {
                ColorConsole.Write(".".Green());
                var pageBreakRange = app.Selection.GoToNext(WdGoToItem.wdGoToPage);
                var links = doc.Range(lastPageEnd, pageBreakRange.End).Hyperlinks;
                if (links.Count > 0)
                {
                    ParseLinks(results, p, links);
                }

                lastPageEnd = pageBreakRange.End;
                links.NAR();
                pageBreakRange.NAR();
                p++;
            }
        }

        private static void ParseLinks(List<Item> results, int p, Hyperlinks links)
        {
            var l = 1;
            foreach (Hyperlink link in links)
            {
                ParseLink(results, p, l, link);
                l++;
            }

            results.Add(new Item(string.Empty, string.Empty, string.Empty));
        }

        private static void ParseLink(List<Item> results, int p, int l, Hyperlink link)
        {
            ColorConsole.Write(".".DarkGray());
            var text = link.TextToDisplay?.Trim();
            if (string.IsNullOrEmpty(text))
            {
                text = HttpUtility.UrlDecode(link.Address).Split('/', StringSplitOptions.RemoveEmptyEntries).LastOrDefault().Replace("-", " ") + "?!";
            }

            results.Add(new Item($"{p}.{l}", text, link.Address.Trim()));
            link.NAR();
        }
    }
}
