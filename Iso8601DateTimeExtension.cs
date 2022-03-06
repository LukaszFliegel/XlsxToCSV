namespace XlsxToCSV
{
    public static class Iso8601DateTimeExtension
    {
        public static string ToIso8601DateOnly(this DateTime value)
        {
            return value.ToString("yyyy-MM-dd");
            // System.DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssK");
        }
        public static string ToIso8601DateTime(this DateTime value)
        {
            return value.ToString("yyyy-MM-ddTHH:mm:ssK");
        }
    }
}
