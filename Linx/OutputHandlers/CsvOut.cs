namespace Linx
{
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;

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
    }
}
