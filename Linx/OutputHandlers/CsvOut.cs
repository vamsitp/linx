namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;

    using ColoredConsole;

    using CsvHelper;
    using CsvHelper.Configuration;

    public class CsvOut : OutputBase
    {
        public override bool Save(List<Item> results, string outputFile)
        {
            if (results?.Count > 0)
            {
                using (var reader = File.CreateText(outputFile))
                {
                    using (var csvWriter = new CsvWriter(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                    {
                        csvWriter.WriteRecords(results);
                    }
                }

                return true;
            }

            return false;
        }

        public override bool Merge(IList<string> outputFiles)
        {
            try
            {
                var mergedOutput = new List<Item>();
                foreach (var file in outputFiles)
                {
                    if (File.Exists(file))
                    {
                        var textReader = new StreamReader(file);
                        using (var csvReader = new CsvReader(textReader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                        {
                            csvReader.Configuration.PrepareHeaderForMatch = (string header, int index) => header.ToLower();
                            var results = csvReader.GetRecords<Item>();
                            mergedOutput.AddRange(results);
                        }

                        mergedOutput.Add(new Item("---", "---", "---"));
                    }
                }

                var first = outputFiles.FirstOrDefault();
                var path = Path.Combine(Path.GetDirectoryName(first), $"{nameof(Linx)}_Merged_[{outputFiles.Count}]{Path.GetExtension(first)}");
                this.Save(mergedOutput, path);
                outputFiles.Add(path);
                return true;
            }
            catch (Exception ex)
            {
                ColorConsole.WriteLine(ex.Message.White().OnRed());
                return false;
            }
        }
    }
}
