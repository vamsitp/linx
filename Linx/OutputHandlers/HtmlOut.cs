namespace Linx
{
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    public class HtmlOut : OutputBase
    {
        public override bool Save(List<Item> results, string outputFile)
        {
            if (results?.Count > 0)
            {
                var html = new StringBuilder($"<html><body style='font-family:Segoe UI'><h3>{this.GetHeader(outputFile)}</h3><hr />");
                foreach (var item in results)
                {
                    var text = string.IsNullOrEmpty(item.Index) ? "<hr />" : $"{item.Index} <a href='{item.Link}'>{item.Text}</a><br />";
                    html.AppendLine(text);
                }

                html.AppendLine("</body></html>");
                File.WriteAllText(outputFile, html.ToString());
                return true;
            }

            return false;
        }
    }
}
