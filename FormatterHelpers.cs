using System;
using System.Drawing;
using Xceed.Document.NET;

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

        public static Paragraph AsH1(this Paragraph paragraph) => paragraph.Heading(HeadingType.Heading1);

        public static Paragraph AsH2(this Paragraph paragraph) => paragraph.Heading(HeadingType.Heading2);

        public static Paragraph AsH3(this Paragraph paragraph) => paragraph.Heading(HeadingType.Heading3);

        public static Paragraph AsH4(this Paragraph paragraph) => paragraph.Heading(HeadingType.Heading4);

        public static Paragraph FormatHealth(this Paragraph paragraph)
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

            return paragraph;
        }
    }
}