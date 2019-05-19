using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace ExcelToCSVMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var configuration = new Configuration
            {
                MergePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\trs-english\\posts\\",
                Worksheets = new[]
                {
                    new WorksheetConfiguration(
                        "Reach",
                        (0, 2),
                        ColumnUnivocity.Index
                    ),
                    new WorksheetConfiguration(
                        "Actions On Post",
                        (1, 1),
                        ColumnUnivocity.Title
                    )
                }
            };

            Console.WriteLine(SheetAnalysis(configuration));

            Console.WriteLine(FileMerger(configuration));

            Console.WriteLine("end");
            Console.ReadLine();
        }

        private static string FileMerger(Configuration configuration)
        {
            return new DirectoryInfo(configuration.MergePath)
                .GetFiles("*.xlsx")
                .SelectMany(ExtractWorksheetsFromFile)
                .Where(sheet => configuration.Worksheets.Any(worksheet => sheet.Name.Contains(worksheet.SheetNameMatch, StringComparison.OrdinalIgnoreCase)))
                .GroupBy(sheet => sheet.Name)
                .Select(grouping => (
                    Name: grouping.Key,
                    Configuration: configuration.Worksheets.Single(worksheet => grouping.Key.Contains(worksheet.SheetNameMatch)),
                    Files: grouping.Select(sheet => sheet.Origin).ToArray()))
                .Select(tuple => MergeIntoExcelPackage(tuple.Name, tuple.Configuration, tuple.Files))
                .Select(table => SaveFile(configuration.MergePath, table))
                .Aggregate(new StringBuilder(), (builder, fileName) => builder.AppendLine($"Saved [{fileName}]"))
                .ToString();
        }

        private static ExcelPackage MergeIntoExcelPackage(String name, WorksheetConfiguration configuration, FileInfo[] files)
        {
            var excelPackage = new ExcelPackage();
            files.Aggregate(excelPackage.Workbook.Worksheets.Add(name), (targetSheet, file) =>
            {
                using (var reader = new ExcelPackage(file))
                using (var source = reader.Workbook.Worksheets[name])
                {
                    var sourceColumns = Enumerable
                        .Range(configuration.StartingPoint.Column, source.Dimension.Columns - configuration.StartingPoint.Column)
                        .Select(columnIndex => (Title: source.Cells[1, columnIndex + 1].Text, Index: columnIndex))
                        .Prepend((Title: "File Name", Index: -1))
                        .ToArray();

                    sourceColumns
                        .Select(configuration.GetColumnName)
                        .Where(columnName => !targetSheet.Cells[1, 1, 1, targetSheet.Dimension?.Columns ?? 1].Any(cell => string.Equals(cell.Text, columnName, StringComparison.OrdinalIgnoreCase)))
                        .Aggregate(targetSheet, (sheet, newColumnName) =>
                        {
                            var newColumnIndex = (sheet.Dimension?.Columns ?? 0) + 1;
                            sheet.InsertColumn(newColumnIndex, 1);
                            sheet.Cells[1, newColumnIndex].Value = newColumnName;
                            return sheet;
                        });

                    var columnsMapping = targetSheet.Cells[1, 1, 1, targetSheet.Dimension.Columns]
                        .Select(header => (
                            SourceColumn: sourceColumns.Single(tuple => configuration.GetColumnName(tuple) == header.Text).Index,
                            TargetColumn: header.Start.Column))
                        .ToArray();

                    Enumerable
                        .Range(configuration.StartingPoint.Row, source.Dimension.Rows - configuration.StartingPoint.Row)
                        .Zip(Enumerable.Range(targetSheet.Dimension.Rows + 1, source.Dimension.Rows - configuration.StartingPoint.Row), (sourceRow, targetRow) => (
                            SourceRow: sourceRow,
                            TargetRow: targetRow))
                        .Select(rowMap =>
                        {
                            return columnsMapping
                                .Aggregate(targetSheet, (sheet, colMap) =>
                                {
                                    if (colMap.SourceColumn == -1)
                                    {
                                        sheet.Cells[rowMap.TargetRow, colMap.TargetColumn].Value = file.Name;
                                        return sheet;
                                    }

                                    source.Cells[rowMap.SourceRow + 1, colMap.SourceColumn + 1].Copy(sheet.Cells[rowMap.TargetRow, colMap.TargetColumn]);
                                    return sheet;
                                });
                        })
                        .ToArray();
                }

                return targetSheet;
            });

            return excelPackage;
        }

        private static string SaveFile(string savePath, ExcelPackage excelPackage)
        {
            var fileInfo = new FileInfo(Path.Combine(savePath, excelPackage.Workbook.Worksheets.First().Name + ".xlsx"));
            excelPackage.SaveAs(fileInfo);
            excelPackage.Dispose();
            return fileInfo.FullName;
        }

        private static string SheetAnalysis(Configuration configuration)
        {
            return new DirectoryInfo(configuration.MergePath)
                .GetFiles("*.xlsx")
                .SelectMany(ExtractWorksheetsFromFile)
                .Where(sheet => configuration.Worksheets.Any(worksheet => sheet.Name.Contains(worksheet.SheetNameMatch, StringComparison.OrdinalIgnoreCase)))
                .OrderBy(sheet => sheet.Name)
                .GroupBy(sheet => sheet.Name)
                .Select(sheets => $"[{sheets.First().Index} - {sheets.All(sheet => sheet.Index == sheets.First().Index)}] [{sheets.Key}]\n\t{string.Join("\n\t", sheets.Select(sheet => $"{sheet.Origin.FullName}, {sheet.NumberOfColumns}"))}")
                .Aggregate(new StringBuilder(), (builder, line) => builder.AppendLine(line))
                .ToString();
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

    internal class Configuration
    {
        public string MergePath { get; set; }
        public WorksheetConfiguration[] Worksheets { get; set; }
    }

    internal class WorksheetConfiguration
    {
        public string SheetNameMatch { get; }
        public (int Column, int Row) StartingPoint { get; }
        private ColumnUnivocity Univocity { get; }

        public WorksheetConfiguration(string sheetNameMatch, (int Column, int Row) startingPoint, ColumnUnivocity univocity) =>
            (SheetNameMatch, StartingPoint, Univocity) = (sheetNameMatch, startingPoint, univocity);

        public string GetColumnName((string Title, int Index) titleAndIndex) =>
            titleAndIndex.Index == -1 || Univocity == ColumnUnivocity.Title ? titleAndIndex.Title : $"{titleAndIndex.Title} [{titleAndIndex.Index}]";
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