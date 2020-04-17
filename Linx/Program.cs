namespace Linx
{
    using System;
    using System.Diagnostics;
    using System.IO;

    using ColoredConsole;

    public class Program
    {
        private static readonly string CurrDir = Environment.CurrentDirectory;
        private static readonly string MyDocsDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        private static readonly string BasePath = Directory.Exists(CurrDir) ? CurrDir : MyDocsDir;

        static void Main(string[] args)
        {
            var file = args[0];
            ColorConsole.WriteLine("input", ": ".Green(), file.DarkGray());
            var results = Path.GetExtension(file).Equals(".pptx") ? DeckEx.ExtractDeckLinks(file) : DocEx.ExtractDocLinks(file);
            results.PrintToConsole();

            var format = OutputFormat.md;
            if (args.Length > 1)
            {
                Enum.TryParse(args[1], out format);
            }

            var outputFile = Path.Combine(BasePath, $"{$"{nameof(Linx)}_{Path.GetFileName(file)}."}{format}");
            if (format == OutputFormat.html)
            {
                results.SaveAsHtml(outputFile);

            }
            else if (format == OutputFormat.csv)
            {
                results.SaveAsCsv(outputFile);
            }
            else
            {
                results.SaveAsMarkdown(outputFile);
            }

            ColorConsole.WriteLine("output", ": ".Green(), outputFile.DarkGray());
            Console.ReadLine();
            Process.Start(new ProcessStartInfo(outputFile) { UseShellExecute = true });
        }
    }

    public class Item
    {
        public Item(string index, string text, string link)
        {
            this.Index = index;
            this.Text = text;
            this.Link = link;
        }

        public string Index { get; set; }
        public string Text { get; set; }
        public string Link { get; set; }
    }

    enum OutputFormat
    {
        md,
        html,
        csv
    }
}
