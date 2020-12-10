using System;
using System.Linq;

namespace MondayMaster
{
    class Program
    {
        private static void Main(string[] args)
        {
            var records = ExcelReader.ReadData().ToList();

            foreach (var r in records)
            {
                Console.WriteLine(r);
            }
            DocGenerator.GenerateDoc(records);
        }

    }
}

