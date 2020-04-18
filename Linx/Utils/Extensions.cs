namespace Linx
{
    using System;
    using System.Runtime.InteropServices;
    using System.Web;

    public static class Extensions
    {
        private const string SafeLinks = "safelinks.protection";

        public static string Sanitize(this string address)
        {
            if (address?.Contains(SafeLinks, StringComparison.OrdinalIgnoreCase) == true)
            {
                return HttpUtility.ParseQueryString(address)[0]?.Trim();
            }

            return address?.Trim();
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
