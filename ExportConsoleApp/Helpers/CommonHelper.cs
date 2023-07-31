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

        public static string? JoinComma<T>(this IEnumerable<T> list, bool noSpace = false)
        {
            if (list == null) return null;
            if (noSpace) return string.Join(",", list);
            return string.Join(", ", list);
        }

    }
}
