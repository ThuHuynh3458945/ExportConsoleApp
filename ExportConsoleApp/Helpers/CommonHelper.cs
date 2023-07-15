namespace ExportConsoleApp.Helpers
{
    public static class CommonHelper
    {
        public static string RemoveComma(this string? content)
        {
            if (string.IsNullOrEmpty(content))
                return "";
            return content.Replace(",", " ");
        }
    }
}
