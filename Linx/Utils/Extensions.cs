namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Web;

    public static class Extensions
    {
        private const string Space = " ";
        private const string SafeLinks = "safelinks.protection";
        private const string WarnSymbol = " ?!";

        private static readonly List<string> Replacements = new List<string> { "here", "this", "link" };
        private static readonly List<string> Cultures = null;

        static Extensions()
        {
            try
            {
                Cultures = CultureInfo.GetCultures(CultureTypes.AllCultures).Select(x => x.Name)?.ToList();
            }
            catch
            {
                // Ignore
            }
        }

        public static string SanitizeLink(this string address)
        {
            if (address?.Contains(SafeLinks, StringComparison.OrdinalIgnoreCase) == true)
            {
                return HttpUtility.ParseQueryString(address)[0]?.Trim();
            }

            return address?.Trim();
        }

        public static string SanitizeText(this string text, string link)
        {
            if (string.IsNullOrWhiteSpace(text) || text.Contains("://") || Replacements.Any(r => text.Equals(r, StringComparison.OrdinalIgnoreCase)))
            {
                var uri = new Uri(HttpUtility.UrlDecode(link));
                var pathAndQuery = uri.PathAndQuery.Split('/', StringSplitOptions.RemoveEmptyEntries);
                if (pathAndQuery?.Any() == true)
                {
                    text = string.Join(" - ", pathAndQuery.Where(p => Cultures == null || (!Cultures.Any(c => c.Equals(p, StringComparison.OrdinalIgnoreCase)) && !p.All(char.IsDigit))).Select(x => x.Replace("-", Space).Replace("_", Space))) + WarnSymbol;
                }
                else
                {
                    var host = uri.Host.Replace("www", Space, StringComparison.OrdinalIgnoreCase).Split('.', StringSplitOptions.RemoveEmptyEntries)?.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Replace("-", Space).Replace("_", Space);
                    text = string.Join(" - ", host) + WarnSymbol;
                }
            }

            return text.ToTitleCase();
        }

        public static string ToTitleCase(this string text)
        {
            var sanitizedText = text.Trim();
            return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(sanitizedText); // char.ToUpper(sanitizedText.First()) + sanitizedText.Substring(1)?.ToLowerInvariant();
        }

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
    }
}
