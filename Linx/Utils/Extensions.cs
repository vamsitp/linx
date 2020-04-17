namespace Linx
{
    using System.Runtime.InteropServices;

    public static class Extensions
    {
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
