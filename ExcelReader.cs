using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MondayMaster
{
    public static class ExcelReader
    {
        private class Comment
        {
            public Comment(IXLRow r)
            {
                DateTime created;

                Id = r.GetValue<string>("A");
                Text = r.GetValue<string>("G");

                IsValid = r.TryGetValue("F", out created);
                Created = created;

                if (string.IsNullOrEmpty(Id)) IsValid = false;
            }

            public bool IsValid { get; }
            public DateTime Created { get; }
            public string Text { get; }
            public string Id { get; }
        }

        private static Dictionary<string, string> ReadComments(XLWorkbook workbook)
        {
            var sheet = workbook.Worksheets.Worksheet("epd-in-flight-updates");
            Console.WriteLine($"Starting to process {sheet.RowCount()} comments");

            var comments = sheet
                .Rows()
                .Select(r => new Comment(r))
                .Where(c => c.IsValid)
                .GroupBy(c => c.Id)
                .Select(commentGroup => commentGroup.OrderByDescending(c => c.Created).First())
                .ToDictionary(c => c.Id, c => c.Text);

            return comments;
        }

        public static IEnumerable<UpdateRecord> ReadData()
        {
            string fileName = $"C:\\Users\\jq\\source\\test\\input.xlsx";
            using (var workbook = new XLWorkbook(fileName))
            {
                var comments = ReadComments(workbook);
                var updates = ReadUpdates(workbook, comments).ToList();

                return updates;
            }
        }

        private static IEnumerable<UpdateRecord> ReadUpdates(XLWorkbook workbook, IDictionary<string, string> comments)
        {
            var currentHeader = string.Empty;

            var sheet = workbook.Worksheets.Worksheet("epd-in-flight");
            Console.WriteLine($"Starting to process {sheet.RowCount()} rows");

            foreach (var row in sheet.Rows().Where(r => !r.IgnoreRow()))
            {
                if (row.IsHeader())
                {
                    currentHeader = row.GetHeader();
                    Console.WriteLine($"Found Header: {currentHeader}");
                }
                else if (!string.IsNullOrEmpty(currentHeader))
                {
                    var record = row.ToUpdateRecord(currentHeader, comments);

                    if (record != null) yield return record;
                }
            }

            Console.WriteLine("Done processing file");
        }
    }
}