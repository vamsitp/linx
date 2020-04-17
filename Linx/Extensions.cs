namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;

    using ColoredConsole;

    using CsvHelper;
    using CsvHelper.Configuration;

    public static class Extensions
    {
        public static void NAR(this object o)
        {
            try
            {
                if (o != null)
                {
                    Marshal.FinalReleaseComObject(o);
                }
            }
            finally
            {
                o = null;
            }
        }

        public static void SaveAsMarkdown(this List<Item> results, string outputFile)
        {
            var md = new StringBuilder();
            foreach (var item in results)
            {
                var text = string.IsNullOrEmpty(item.Index) ? "---" : $"{item.Index} [{item.Text}]({item.Link}){Environment.NewLine}";
                md.AppendLine(text);
            }

            File.WriteAllText(outputFile, md.ToString());
        }

        public static void SaveAsHtml(this List<Item> results, string outputFile)
        {
            var html = new StringBuilder("<html><body style='font-family:Segoe UI'>");
            foreach (var item in results)
            {
                var text = string.IsNullOrEmpty(item.Index) ? "<hr />" : $"{item.Index} <a href='{item.Link}'>{item.Text}</a><br />";
                html.AppendLine(text);
            }

            html.AppendLine("</body></html>");
            File.WriteAllText(outputFile, html.ToString());
        }

        public static void SaveAsCsv(this List<Item> results, string outputFile)
        {
            using (var reader = File.CreateText(outputFile))
            {
                using (var csvWriter = new CsvWriter(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                {
                    csvWriter.WriteRecords(results);
                }
            }
        }

        public static void PrintToConsole(this List<Item> results)
        {
            foreach (var item in results)
            {
                var text = string.IsNullOrEmpty(item.Index) ? new[] { "-----".Green(), Environment.NewLine } : new[] { item.Index.Green(), " ", item.Text, Environment.NewLine, item.Link.Blue(), Environment.NewLine };
                ColorConsole.WriteLine(text);
            }
        }
    }
}
