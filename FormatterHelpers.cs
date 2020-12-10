using System;
using System.Drawing;

namespace MondayMaster
{
    public static class FormatterHelpers
    {
        public static string FormatDate(this DateTime? dt)
        {
            return dt.HasValue
                ? dt.Value.ToShortDateString()
                : "";

        }
        public static void FormatHealth(this Xceed.Document.NET.Paragraph paragraph)
        {
            switch (paragraph.Text)
            {
                case "At Risk":
                    paragraph.Color(Color.Orange);
                    break;
                case "Unhealthy":
                    paragraph.Color(Color.Red);
                    break;
                case "Healthy":
                    paragraph.Color(Color.Green);
                    break;
                default:
                    break;
            }
        }
    }
}