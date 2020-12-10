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

            var doc = DocX.Create(fileName);

            doc.InsertParagraph("Hello Word");

            AddTable(doc, records);

            doc.Save();

            Process.Start("WINWORD.EXE", fileName);
        }


        private static void AddTable(DocX doc, List<UpdateRecord> records)
        {
            Table t = doc.AddTable(records.Count(), 4);
            t.Alignment = Alignment.center;
            t.Design = TableDesign.ColorfulList;

            //Fill cells by adding text.  
            t.Rows[0].Cells[0].Paragraphs.First().Append("Stage");
            t.Rows[0].Cells[1].Paragraphs.First().Append("Name");
            t.Rows[0].Cells[2].Paragraphs.First().Append("Health");
            t.Rows[0].Cells[3].Paragraphs.First().Append("Update");

            for (int i = 1; i < records.Count; i++)
            {
                var record = records[i];

                t.Rows[i].Cells[0].Paragraphs.First().Append(record.Header);
                t.Rows[i].Cells[1].Paragraphs.First().Append(record.Name);
                t.Rows[i].Cells[2].Paragraphs.First().Append(record.Health);
                t.Rows[i].Cells[3].Paragraphs.First().Append(record.Comment);
            }

            doc.InsertTable(t);
        }
    }
}

