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
                InsertPreamble(doc);
                InsertTable(doc, records);
                InsertMetrics(doc);

                doc.Save();
            }

            Process.Start("WINWORD.EXE", fileName);
        }

        private static void InsertMetrics(DocX doc)
        {
            doc.InsertParagraph("Key Metrics").AsH2();
            doc.InsertParagraph("See full metrics over at the Integrations adoption dashboard.");

            doc.InsertParagraph("");

            doc.InsertParagraph("*Adoption is calculated by total installs divided by total customers reporting the vendor in Salesforce. Since Salesforce data is manually updated, there can be some data quality issues.");
        }

        private static void InsertPreamble(DocX doc)
        {
            doc.InsertParagraph(DateTime.Now.ToString("MM/d")).AsH1();
            doc.InsertParagraph("Hey there everyone!\r\n");
            doc.InsertParagraph("See below for the Integrations Team weekly update! Let us know if you have any questions!\r\n");
            doc.InsertParagraph("All the best, \r\nThe Integrations Team\r\n");

            doc.InsertParagraph("Week in Review").AsH2();

            doc.InsertParagraph("Integration Core").AsH3();
            doc.InsertParagraph("Highlights").AsH4();
            doc.AddBulletedList();

            doc.InsertParagraph("Lowlights").AsH4();
            doc.AddBulletedList();
            doc.InsertParagraph("What is preventing your team from doing their best work?").AsH4();

            doc.InsertParagraph("Integration Apps").AsH3();
            doc.InsertParagraph("Highlights").AsH4();
            doc.AddBulletedList();
            doc.InsertParagraph("Lowlights").AsH4();
            doc.AddBulletedList();
            doc.InsertParagraph("What is preventing your team from doing their best work?").AsH4();

            doc.InsertParagraph("Weekly Update").AsH2();
        }

        private static void InsertTable(DocX doc, List<UpdateRecord> records)
        {
            var groups = records.GroupBy(r => r.Header);

            foreach (var group in groups)
            {
                var header = doc.InsertParagraph(group.Key);
                header.Heading(HeadingType.Heading2);
                AddTable(doc, group.ToList());
            }
        }

        private static void AddTable(DocX doc, List<UpdateRecord> records)
        {
            //foreach (TableDesign td in Enum.GetValues(typeof(TableDesign)))
            //{
            //    var header = doc.InsertParagraph(td.ToString());
            //    header.Heading(HeadingType.Heading2);
            //    AddTable(doc, records, td);
            //}

            AddTable(doc, records, TableDesign.LightGridAccent1);
        }

        private static void AddTable(DocX doc, List<UpdateRecord> records, TableDesign td)
        {
            Table t = doc.AddTable(1, 4);
            t.Alignment = Alignment.left;
            t.Design = td;

            t.SetColumnWidth(0, 90);
            t.SetColumnWidth(1, 70);
            t.SetColumnWidth(2, 70);
            t.SetColumnWidth(3, 310);

            int c = 0;
            t.Rows[0].Cells[c++].Paragraphs.First().Append("Name");
            t.Rows[0].Cells[c++].Paragraphs.First().Append("Health");
            t.Rows[0].Cells[c++].Paragraphs.First().Append("Exit Date\r\nOriginal\r\nCurrent");
            t.Rows[0].Cells[c++].Paragraphs.First().Append("Comment");

            foreach (var record in records)
            {
                var row = t.InsertRow();
                c = 0;

                row.Cells[c++].Paragraphs.First().Append(record.Name);
                row.Cells[c++].Paragraphs.First().Append(record.Health).FormatHealth();
                row.Cells[c++].Paragraphs.First()
                    .Append(record.ExitDateOriginal.FormatDate())
                    .Append(System.Environment.NewLine)
                    .Append(System.Environment.NewLine)
                    .Append(record.ExitDateCurrent.FormatDate());

                row.Cells[c++].Paragraphs.First().Append(record.Comment);
            }

            doc.InsertTable(t);
        }
    }
}