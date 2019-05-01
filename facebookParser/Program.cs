using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using ExcelDataReader;

namespace facebookParser
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var excelContainingDirectory = new DirectoryInfo("c:\\Users\\GCATBV\\Downloads\\testdata\\");
            var sheetNamesList = excelContainingDirectory
                .GetFiles("*.xlsx")
                .Select(file =>
                {
                    using (var stream = file.OpenRead())
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    using (var dataSet = reader.AsDataSet())
                    {
                        var tableNames = new List<string>();
                        for (int i = 0; i < dataSet.Tables.Count; i++)
                        {
                            tableNames.Add(dataSet.Tables[i].TableName);
                        }
                        return tableNames.ToArray();
                    }
                })
                .ToArray();

            foreach (var sheetNames in sheetNamesList)
            {
                Console.WriteLine(string.Join(",", sheetNames) + ".");
            }
            Console.WriteLine("union:");
            var union = sheetNamesList.Aggregate(new List<string>(), (list, strings) => list.Union(strings).ToList());
            Console.WriteLine(string.Join(",", union) + ".");
            Console.WriteLine("differences:");
            foreach (var sheetNames in sheetNamesList)
            {
                Console.WriteLine(string.Join(",", union.Except(sheetNames)) + ".");
            }
            Console.WriteLine("end");

            Console.ReadLine();
        }
    }
}