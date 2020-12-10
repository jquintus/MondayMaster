using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace MondayMaster
{
    public static class DocGenerator
    {
        public static void GenerateDoc(List<UpdateRecord> records)
        {
            var stamp = DateTime.Now.ToString("yy-MM-dd-HH-mm-ss");
            string fileName = $"C:\\Users\\jq\\source\\test\\out_{stamp}.docx";

            using (var doc = DocX.Create(fileName))
            {
                var groups = records.GroupBy(r => r.Header);

                foreach (var group in groups)
                {
                    var header = doc.InsertParagraph(group.Key);
                    header.Heading(HeadingType.Heading2);
                    AddTable(doc, group.ToList());
                }

                doc.Save();
            }

            Process.Start("WINWORD.EXE", fileName);
        }

        private static void AddTable(DocX doc, List<UpdateRecord> records)
        {
            Table t = doc.AddTable(records.Count() + 1, 5);
            t.Alignment = Alignment.left;
            t.Design = TableDesign.ColorfulListAccent2;

            Console.WriteLine(t.GetColumnWidth(0));
            Console.WriteLine(t.GetColumnWidth(1));
            Console.WriteLine(t.GetColumnWidth(2));
            Console.WriteLine(t.GetColumnWidth(3));
            Console.WriteLine(t.GetColumnWidth(4));

            t.SetColumnWidth(0, 90);
            t.SetColumnWidth(1, 70);
            t.SetColumnWidth(2, 70);
            t.SetColumnWidth(3, 70);
            t.SetColumnWidth(4, 240);

            int c = 0;
            t.Rows[0].Cells[c++].Paragraphs.First().Append("Name");
            t.Rows[0].Cells[c++].Paragraphs.First().Append("Health");
            t.Rows[0].Cells[c++].Paragraphs.First().Append("Exit Date (original)");
            t.Rows[0].Cells[c++].Paragraphs.First().Append("Exit Date (current)");
            t.Rows[0].Cells[c++].Paragraphs.First().Append("Comment");

            for (int i = 1; i <= records.Count; i++)
            {
                var record = records[i-1];
                c = 0;

                t.Rows[i].Cells[c++].Paragraphs.First().Append(record.Name);
                t.Rows[i].Cells[c++].Paragraphs.First().Append(record.Health).FormatHealth();
                t.Rows[i].Cells[c++].Paragraphs.First().Append(record.ExitDateOriginal.FormatDate());
                t.Rows[i].Cells[c++].Paragraphs.First().Append(record.ExitDateCurrent.FormatDate());

                t.Rows[i].Cells[c++].Paragraphs.First().Append(record.Comment);

            }

            doc.InsertTable(t);
        }
    }
}