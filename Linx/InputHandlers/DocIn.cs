namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using ColoredConsole;

    using Microsoft.Office.Core;
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
            ColorConsole.Write(p.ToString().Green());
            Range pageBreakRange = null;
            try
            {
                pageBreakRange = app.Selection.GoToNext(WdGoToItem.wdGoToPage);
                var links = doc.Range(lastPageEnd, pageBreakRange.End).Hyperlinks;
                if (links.Count > 0)
                {
                    ParseLinks(results, p, links.Cast<Hyperlink>()?.Select(link =>
                    {
                        var text = link.Type == MsoHyperlinkType.msoHyperlinkRange ? link.TextToDisplay?.Trim() : string.Empty;
                        var item = new Item(string.Empty, text, link.Address);
                        link.NAR();
                        return item;
                    })?.ToList());
                    links.NAR();
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
    }
}
