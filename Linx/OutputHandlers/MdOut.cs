namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    public class MdOut : OutputBase
    {
        public override bool Save(List<Item> results, string outputFile)
        {
            if (results?.Count > 0)
            {
                var md = new StringBuilder($"### {this.GetHeader(outputFile)}{Environment.NewLine}---{Environment.NewLine}{Environment.NewLine}");
                foreach (var item in results)
                {
                    var text = string.IsNullOrEmpty(item.Index) ? "---" : $"{item.Index} [{item.Text}]({item.Link}){Environment.NewLine}";
                    md.AppendLine(text);
                }

                File.WriteAllText(outputFile, md.ToString());
                return true;
            }

            return false;
        }
    }
}
