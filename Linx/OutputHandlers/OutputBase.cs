namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;

    using ColoredConsole;

    public interface IOutput
    {
        bool Save(List<Item> results, string outputFile);
        bool Merge(IList<string> outputFiles);
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

        public virtual bool Merge(IList<string> outputFiles)
        {
            try
            {
                var mergedOutput = new StringBuilder();
                foreach (var file in outputFiles)
                {
                    mergedOutput.Append(File.ReadAllText(file));
                    mergedOutput.AppendLine();
                }

                var first = outputFiles.FirstOrDefault();
                var path = Path.Combine(Path.GetDirectoryName(first), $"{nameof(Linx)}_Merged_[{outputFiles.Count}]{Path.GetExtension(first)}");
                File.WriteAllText(path, mergedOutput.ToString());
                outputFiles.Add(path);
                return true;
            }
            catch (Exception ex)
            {
                ColorConsole.WriteLine(ex.Message.White().OnRed());
                return false;
            }
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
