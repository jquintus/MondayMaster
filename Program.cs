using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace MondayMaster
{
    class Program
    {
        private static void Main(string[] args)
        {
            var stamp = DateTime.Now.ToString("yy-MM-dd-HH-mm-ss");
            string fileName = $"C:\\Users\\jq\\source\\test\\exampleWord_{stamp}.docx";

            var doc = DocX.Create(fileName);

            doc.InsertParagraph("Hello Word");


            AddTable(doc);

            doc.Save();

            Process.Start("WINWORD.EXE", fileName);
        }

        private static void AddTable(DocX doc)
        {
            Table t = doc.AddTable(5, 4);
            t.Alignment = Alignment.center;
            t.Design = TableDesign.ColorfulList;

            //Fill cells by adding text.  
            t.Rows[0].Cells[0].Paragraphs.First().Append("AA");
            t.Rows[0].Cells[1].Paragraphs.First().Append("BB");
            t.Rows[0].Cells[2].Paragraphs.First().Append("CC");
            t.Rows[0].Cells[3].Paragraphs.First().Append("DD");

            for (int i = 1; i < 5; i++)
            {
                t.Rows[i].Cells[0].Paragraphs.First().Append("EE");
                t.Rows[i].Cells[1].Paragraphs.First().Append("FF");
                t.Rows[i].Cells[2].Paragraphs.First().Append("GG");
                t.Rows[i].Cells[3].Paragraphs.First().Append("HH");
            }

            doc.InsertTable(t);
        }
    }
}
