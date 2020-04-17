namespace Linx
{
    using System;
    using System.Collections.Generic;

    using ColoredConsole;

    public class ConsoleOut : OutputBase
    {
        public override bool Save(List<Item> results, string outputFile)
        {
            if (results?.Count > 0)
            {
                foreach (var item in results)
                {
                    var text = string.IsNullOrEmpty(item.Index) ? new[] { "-----".DarkGray(), Environment.NewLine } : new[] { item.Index.Green(), " ", item.Text.DarkGray(), Environment.NewLine, item.Link, Environment.NewLine };
                    ColorConsole.WriteLine(text);
                }

                return true;
            }

            return false;
        }
    }
}
