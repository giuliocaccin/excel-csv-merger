using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;

namespace facebookParser
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var pathToExcels = "c:\\Users\\GCATBV\\Downloads\\testdata\\";

//            SheetAnalysis(pathToExcels);

            FileMerger(pathToExcels);

            Console.ReadLine();
        }

        private static void FileMerger(string pathToExcels)
        {
            var sheetConfiguration = new[]
            {
                new SheetConfiguration
                {
                    SheetNameMatch = "Key Metrics",
                    StartingPoint = (0, 2)
                },
                new SheetConfiguration
                {
                    SheetNameMatch = "Daily",
                    StartingPoint = (1, 1)
                }
            };

            var sheetNamesMatch = sheetConfiguration.Select(configuration => configuration.SheetNameMatch).ToArray();

            var newFiles = new DirectoryInfo(pathToExcels)
                .GetFiles("*.xlsx")
                .SelectMany(ExtractSheet)
                .Where(sheet => sheetNamesMatch.Any(_ => sheet.Name.Contains(_, StringComparison.OrdinalIgnoreCase)))
                .GroupBy(sheet => sheet.Name)
                .Select(grouping => (
                    Name: grouping.Key,
                    Configuration: sheetConfiguration.Single(configuration => grouping.Key.Contains(configuration.SheetNameMatch)),
                    Files: grouping.Select(sheet => sheet.FileName).ToArray()))
                .Select(MergeIntoDataTable)
                .Select(DataTableToCsv)
                .ToArray();

            foreach (var (name, content) in newFiles)
            {
                File.WriteAllText($"{pathToExcels}{name}.csv", content, Encoding.UTF8);
            }
        }

        private static DataTable MergeIntoDataTable((String Name, SheetConfiguration Configuration, string[] Files) configuration)
        {
            var (name, sheetConfiguration, files) = configuration;
            var target = new DataTable(name);
            target.Columns.Add("fileName");
            foreach (var path in files)
            {
                using (var stream = File.OpenRead(path))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                using (var dataSet = reader.AsDataSet())
                using (var source = dataSet.Tables[name])
                {
                    target.BeginLoadData();
                    foreach (var sourceColumn in source.Columns.Cast<DataColumn>().Skip(sheetConfiguration.StartingPoint.X))
                    {
                        var sourceColumnName = source.Rows[0].ItemArray[sourceColumn.Ordinal].ToString();
                        if (!target.Columns.Contains(sourceColumnName))
                            target.Columns.Add(sourceColumnName);
                    }

                    foreach (var sourceRow in source.AsEnumerable().Skip(sheetConfiguration.StartingPoint.Y))
                    {
                        var newTargetRow = target.NewRow();
                        newTargetRow.SetField("fileName", path);
                        foreach (var sourceColumn in source.Columns.Cast<DataColumn>().Skip(sheetConfiguration.StartingPoint.X))
                        {
                            var sourceColumnName = source.Rows[0].ItemArray[sourceColumn.Ordinal].ToString();
                            newTargetRow.SetField(sourceColumnName, sourceRow.ItemArray[sourceColumn.Ordinal]);
                        }

                        target.Rows.Add(newTargetRow);
                    }

                    target.EndLoadData();
                }
            }

            return target;
        }

        private static (string Name, string Content) DataTableToCsv(DataTable dataTable) =>
            (Name: dataTable.TableName,
                Content: dataTable.AsEnumerable()
                    .Aggregate(
                        new StringBuilder(string.Join(",", dataTable.Columns.Cast<DataColumn>().Select(column => $"\"{column.ColumnName}\""))),
                        (builder, row) =>
                        {
                            builder.AppendLine(string.Join(",", row.ItemArray.Select(item => $"\"{item}\"")));
                            return builder;
                        })
                    .ToString());

        private static void SheetAnalysis(string pathToExcels)
        {
            var sheetNamesList = new DirectoryInfo(pathToExcels)
                .GetFiles("*.xlsx")
                .SelectMany(ExtractSheet)
                .OrderBy(sheet => sheet.Name)
                .GroupBy(sheet => sheet.Name)
                .Select(sheets => $"[{sheets.First().Position} - {sheets.All(sheet => sheet.Position == sheets.First().Position)}] [{sheets.Key}]\n\t{string.Join("\n\t", sheets.Select(sheet => $"{sheet.FileName}, {sheet.NumberOfColumns}"))}")
                .ToArray();
            foreach (var sheet in sheetNamesList)
            {
                Console.WriteLine(sheet);
            }
        }

        private static IEnumerable<Sheet> ExtractSheet(FileInfo file)
        {
            using (var stream = file.OpenRead())
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            using (var dataSet = reader.AsDataSet())
            {
                for (int i = 0; i < dataSet.Tables.Count; i++)
                {
                    var dataTable = dataSet.Tables[i];
                    yield return new Sheet(
                        i,
                        file.FullName,
                        dataTable.TableName,
                        dataTable.Columns.Count);
                }
            }
        }
    }

    internal class SheetConfiguration
    {
        public string SheetNameMatch { get; set; }
        public (int X, int Y) StartingPoint { get; set; }
    }

    class Sheet
    {
        public int Position { get; }
        public string FileName { get; }
        public string Name { get; }
        public int NumberOfColumns { get; }

        public Sheet(int position, string fileName, string name, int numberOfColumns)
        {
            Position = position;
            FileName = fileName;
            Name = name;
            NumberOfColumns = numberOfColumns;
        }
    }
}