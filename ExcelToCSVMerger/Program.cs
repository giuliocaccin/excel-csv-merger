using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace ExcelToCSVMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var pathToExcels = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\trs-english\\";

            //SheetAnalysis(pathToExcels);

            FileMerger(pathToExcels);

            Console.WriteLine("end");
            Console.ReadLine();
        }

        private static void FileMerger(string pathToExcels)
        {
            var sheetConfiguration = new[]
            {
                new WorksheetConfiguration(
                    "Key Metrics",
                    (0, 2),
                    ColumnUnivocity.Index
                ),
                new WorksheetConfiguration(
                    "Lifetime Likes By Country",
                    (1, 1),
                    ColumnUnivocity.Title
                )
            };

            var validWorksheetNames = sheetConfiguration.Select(configuration => configuration.SheetNameMatch).ToArray();

            var newFiles = new DirectoryInfo(pathToExcels)
                .GetFiles("*.xlsx")
                .SelectMany(ExtractWorksheetsFromFile)
                .Where(sheet => validWorksheetNames.Any(name => sheet.Name.Contains(name, StringComparison.OrdinalIgnoreCase)))
                .GroupBy(sheet => sheet.Name)
                .Select(grouping => (
                    Name: grouping.Key,
                    Configuration: sheetConfiguration.Single(configuration => grouping.Key.Contains(configuration.SheetNameMatch)),
                    Files: grouping.Select(sheet => sheet.Origin).ToArray()))
                .Select(tuple => MergeIntoDataTable(tuple.Name, tuple.Configuration, tuple.Files))
                .Select(DataTableToCsv)
                .ToArray();

            foreach (var (name, content) in newFiles)
            {
                File.WriteAllText($"{pathToExcels}{name}.csv", content, Encoding.UTF8);
            }
        }

        private static DataTable MergeIntoDataTable(String name, WorksheetConfiguration configuration, FileInfo[] files)
        {
            const string firstColumnName = "fileName";
            var target = new DataTable(name);
            target.Columns.Add(firstColumnName);
            foreach (var file in files)
            {
                using (var reader = new ExcelPackage(file))
                using (var source = reader.Workbook.Worksheets[name])
                {
                    target.BeginLoadData();

                    var titleAndIndexList = Enumerable
                        .Range(configuration.StartingPoint.Column, source.Dimension.Columns - configuration.StartingPoint.Column)
                        .Select(columnIndex => (Title: source.Cells[1, columnIndex + 1].Text, Index: columnIndex))
                        .ToArray();

                    target.Columns.AddRange(
                        titleAndIndexList
                            .Where(titleAndIndex => !target.Columns.Contains(configuration.GetColumnName(titleAndIndex)))
                            .Select(titleAndIndex => new DataColumn(configuration.GetColumnName(titleAndIndex)))
                            .ToArray()
                    );

                    var rowsToAdd = Enumerable
                        .Range(configuration.StartingPoint.Row, source.Dimension.Rows - configuration.StartingPoint.Row)
                        .Select(rowIndex =>
                        {
                            var newTargetRow = target.NewRow();
                            newTargetRow.SetField(firstColumnName, file.Name);
                            foreach (var titleAndIndex in titleAndIndexList)
                                newTargetRow.SetField(configuration.GetColumnName(titleAndIndex), source.Cells[rowIndex + 1, titleAndIndex.Index + 1].Text);
                            return newTargetRow;
                        })
                        .ToArray();

                    foreach (var row in rowsToAdd)
                    {
                        target.Rows.Add(row);
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
                        new StringBuilder(string.Join(",", dataTable.Columns.Cast<DataColumn>().Select(column => $"\"{column.ColumnName}\"")) + Environment.NewLine),
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
                .SelectMany(ExtractWorksheetsFromFile)
                .OrderBy(sheet => sheet.Name)
                .GroupBy(sheet => sheet.Name)
                .Select(sheets => $"[{sheets.First().Index} - {sheets.All(sheet => sheet.Index == sheets.First().Index)}] [{sheets.Key}]\n\t{string.Join("\n\t", sheets.Select(sheet => $"{sheet.Origin.FullName}, {sheet.NumberOfColumns}"))}")
                .ToArray();
            foreach (var sheet in sheetNamesList)
            {
                Console.WriteLine(sheet);
            }
        }

        private static IReadOnlyList<Worksheet> ExtractWorksheetsFromFile(FileInfo file)
        {
            using (var excelFile = new ExcelPackage(file))
                return excelFile.Workbook.Worksheets
                    .Select(worksheet =>
                        new Worksheet(worksheet.Index, file, worksheet.Name, worksheet.Dimension.Columns))
                    .ToArray();
        }
    }

    internal class WorksheetConfiguration
    {
        public string SheetNameMatch { get; }
        public (int Column, int Row) StartingPoint { get; }
        private ColumnUnivocity Univocity { get; }

        public WorksheetConfiguration(string sheetNameMatch, (int Column, int Row) startingPoint, ColumnUnivocity univocity) =>
            (SheetNameMatch, StartingPoint, Univocity) = (sheetNameMatch, startingPoint, univocity);

        public string GetColumnName((string Title, int Index) titleAndIndex) =>
            Univocity == ColumnUnivocity.Title ? titleAndIndex.Title : $"{titleAndIndex.Title} [{titleAndIndex.Index}]";
    }

    internal enum ColumnUnivocity
    {
        Index,
        Title
    }

    internal class Worksheet
    {
        public int Index { get; }
        public FileInfo Origin { get; }
        public string Name { get; }
        public int NumberOfColumns { get; }

        public Worksheet(int index, FileInfo origin, string name, int numberOfColumns) =>
            (Index, Origin, Name, NumberOfColumns) = (index, origin, name, numberOfColumns);
    }
}