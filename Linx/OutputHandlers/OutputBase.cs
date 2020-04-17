namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    public interface IOutput
    {
        bool Save(List<Item> results, string outputFile);
    }

    public abstract class OutputBase : IOutput
    {
        private static readonly Dictionary<OutputFormat, IOutput> Outputs = new Dictionary<OutputFormat, IOutput>
        {
            { OutputFormat.md, new MdOut() },
            { OutputFormat.html, new HtmlOut()},
            { OutputFormat.csv, new CsvOut() }
        };

        public static IOutput GetInstance(OutputFormat format)
        {
            return Outputs[format];
        }

        public virtual bool Save(List<Item> results, string outputFile)
        {
            throw new NotImplementedException();
        }

        protected string GetHeader(string outputFile)
        {
            var header = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(outputFile)).Replace("Linx_", "Linx > ").Replace("-", " ").Replace("_", " ");
            return header;
        }
    }

    public enum OutputFormat
    {
        md,
        html,
        csv
    }
}
