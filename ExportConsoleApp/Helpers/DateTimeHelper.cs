namespace ExportConsoleApp.Heplers
{
    public static class DateTimeHelper
    {
        public static DateTime ConvertToEstTime(DateTime utcDateTime)
        {
            var mountain = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
            var estDateTime = TimeZoneInfo.ConvertTimeFromUtc(utcDateTime, mountain);
            return estDateTime;
        }
    }
}
