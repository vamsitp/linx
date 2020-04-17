namespace Linx
{
    using System;
    using System.Collections.Generic;

    using ColoredConsole;

    using Microsoft.Office.Interop.Word;

    using Range = Microsoft.Office.Interop.Word.Range;

    public class DocIn : InputBase
    {
        public override List<Item> ExtractLinks(object file)
        {
            List<Item> results = null;
            Application app = null;
            Document doc = null;

            try
            {
                app = new Application();
                object missing = Type.Missing;
                doc = app.Documents.Open(
                                ref file,
                                ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);
                results = ParsePages(app, doc);
            }
            catch (Exception ex)
            {
                ColorConsole.WriteLine(ex.Message.White().OnRed());
            }
            finally
            {
                doc?.Close();
                doc?.NAR();
                app?.Quit();
                app.NAR();
            }

            ColorConsole.WriteLine(Environment.NewLine);
            return results;
        }

        private List<Item> ParsePages(Application app, Document doc)
        {
            var results = new List<Item>();
            var p = 1;
            var pageCount = doc.ComputeStatistics(WdStatistic.wdStatisticPages);
            var lastPageEnd = 0;
            for (long i = 0; i < pageCount; i++)
            {
                lastPageEnd = ParsePage(results, app, doc, p, lastPageEnd);
                p++;
            }

            return results;
        }

        private int ParsePage(List<Item> results, Application app, Document doc, int p, int lastPageEnd)
        {
            ColorConsole.Write(".".White());
            Range pageBreakRange = null;
            try
            {
                pageBreakRange = app.Selection.GoToNext(WdGoToItem.wdGoToPage);
                var links = doc.Range(lastPageEnd, pageBreakRange.End).Hyperlinks;
                if (links.Count > 0)
                {
                    ParseLinks(results, p, links);
                }

                lastPageEnd = pageBreakRange.End;
            }
            catch (Exception ex)
            {
                ColorConsole.WriteLine(ex.Message.White().OnRed());
                lastPageEnd++; // TODO: Check
            }
            finally
            {
                pageBreakRange.NAR();
            }

            return lastPageEnd;
        }

        private void ParseLinks(List<Item> results, int s, Hyperlinks links)
        {
            var l = 1;
            foreach (Hyperlink link in links)
            {
                var text = string.Empty;
                try
                {
                    text = link.TextToDisplay?.Trim();
                }
                catch
                {
                    ColorConsole.Write(" ".OnRed());
                }

                ParseLink(results, s, l, text, link.Address);
                link.NAR();
                l++;
            }

            results.Add(new Item(string.Empty, string.Empty, string.Empty));
            links.NAR();
        }
    }
}
