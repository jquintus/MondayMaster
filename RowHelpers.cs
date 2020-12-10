using ClosedXML.Excel;
using System;
using System.Linq;

namespace MondayMaster
{
    public static class RowHelpers
    {
        public static DateTime? TryToDateTime(this string date)
        {
            if (DateTime.TryParse(date, out DateTime result))
            {
                return result;
            }
            else
            {
                return null;
            }
        }

        public static bool IgnoreRow(this IXLRow row)
        {
            if (row == null) return true;
            return row.Cells().All(c => string.IsNullOrWhiteSpace(c.GetValue<string>()));
        }

        public static T GetValue<T>(this IXLRow row, string columnLetter) => row.Cell(columnLetter).GetValue<T>();

        public static bool TryGetValue<T>(this IXLRow row, string columnLetter, out T value)
        {
            return row.Cell(columnLetter)
                      .TryGetValue<T>(out value);
        }

        public static UpdateRecord ToUpdateRecord(this IXLRow row, string header, System.Collections.Generic.IDictionary<string, string> comments)
        {
            var name = row.GetValue<string>("A");
            var health = row.GetValue<string>("C");
            var teams = row.GetValue<string>("D");
            var pm = row.GetValue<string>("E");
            var leadEng = row.GetValue<string>("F");
            var exitOriginalStr = row.GetValue<string>("I");
            var exitCurrentStr = row.GetValue<string>("J");
            var id = row.GetValue<string>("L");

            if (name == "Name" && health == "Health" && teams == "Teams") return null;
            if (string.IsNullOrEmpty(name)) return null;

            comments.TryGetValue(id, out string comment);

            return new UpdateRecord
            {
                Id = id,
                Name = name,
                Health = health,
                Comment = comment,
                Header = header,
                ProductManager = pm,
                LeadEng = leadEng,
                ExitDateCurrent = exitCurrentStr.TryToDateTime(),
                ExitDateOriginal = exitOriginalStr.TryToDateTime(),
            };
        }

        public static bool IsHeader(this IXLRow row)
        {
            if (row == null) return false;

            string[] statuses = new string[] {
                "Execution",
                "Ready",
                "Requirements",
                "Discovery",
            };
            string cell1 = row.Cell(1).GetValue<string>();
            string cell2 = row.Cell(2).GetValue<string>();

            return statuses.Contains(cell1) && string.IsNullOrEmpty(cell2);
        }

        public static string GetHeader(this IXLRow row) => row.Cell(1).GetValue<string>();
    }
}