using CommandLine;
using Mustache;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

namespace Populate
{
    public static class Program
    {
        public static async Task<int> Main(string[] args)
        {
            var parserResult = Parser.Default.ParseArguments<CommandLineOptions>(args);

            await parserResult.WithParsedAsync(RunAsync);

            return Environment.ExitCode;
        }

        private static async Task RunAsync(CommandLineOptions options)
        {
            if (!File.Exists(options.TemplateFile))
            {
                Console.WriteLine($"The template file does not exist: {options.TemplateFile}");

                return;
            }

            if (!File.Exists(options.ExcelFile))
            {
                Console.WriteLine($"The Excel file does not exist: {options.ExcelFile}");

                return;
            }

            if (!Directory.Exists(options.OutputDirectory))
            {
                Console.WriteLine($"The output directory does not exist: {options.OutputDirectory}");

                return;
            }

            var template = await File.ReadAllTextAsync(options.TemplateFile);

            if (string.IsNullOrEmpty(template))
            {
                Console.WriteLine($"The template file is empty: {options.TemplateFile}");

                return;
            }

            var records = await ReadExcelAsync(options.ExcelFile, CultureInfo.CurrentUICulture);

            Console.WriteLine($"{records.Count} records found in Excel file.");

            var formatCompiler =
                new FormatCompiler
                {
                    AreExtensionTagsAllowed = true,
                    RemoveNewLines = true
                };

            var generator = formatCompiler.Compile(template);

            var index = 0;

            foreach (var record in records)
            {
                index++;

                var content = generator.Render(CultureInfo.CurrentUICulture, record);

                var outputFileName = $"output-{index}.{options.Extension}";

                if (!string.IsNullOrEmpty(options.OutputFilenameColumn) &&
                    record.TryGetValue(options.OutputFilenameColumn, out var alternativeFileName) &&
                    !string.IsNullOrEmpty(alternativeFileName))
                {
                    outputFileName = $"{alternativeFileName}.{options.Extension}";
                }

                var outputFilePath = Path.Combine(options.OutputDirectory, outputFileName);

                Console.WriteLine($"Saving {outputFilePath}");

                await File.WriteAllTextAsync(outputFilePath, content);
            }

            Console.WriteLine($"Processing {options.ExcelFile} completed.");
        }

        private static async Task<IList<IDictionary<string, string>>> ReadExcelAsync(string filePath, CultureInfo cultureInfo)
        {
            await using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);

            using var package = new ExcelPackage(stream);

            package.Compatibility.IsWorksheets1Based = true;

            var worksheet = package.Workbook.Worksheets[1];

            var worksheetDimension = worksheet.Dimension;

            if (worksheetDimension == null)
                throw new Exception("Worksheet is empty.");

            var columnNoToHeader = new Dictionary<int, string>();

            for (var columnNo = 1; columnNo <= worksheetDimension.Columns; columnNo++)
            {
                var cell = worksheet.Cells[1, columnNo];

                if (cell.Value == null ||
                    !(cell.Value is string stringValue) ||
                    string.IsNullOrWhiteSpace(stringValue))
                    throw new Exception(
                        $"Invalid header name on row: 1, column: {columnNo}, cell address: {cell.Address}");

                columnNoToHeader.Add(columnNo, stringValue);
            }

            if (columnNoToHeader.Count == 0)
                throw new Exception("Worksheet is empty.");

            var result = new List<IDictionary<string, string>>();

            for (var rowNo = 2; rowNo <= worksheetDimension.Rows; rowNo++)
            {
                var row = new Dictionary<string, string>();

                for (var columnNo = 1; columnNo <= worksheetDimension.Columns; columnNo++)
                {
                    var cell = worksheet.Cells[rowNo, columnNo];

                    if (!columnNoToHeader.TryGetValue(columnNo, out var header))
                        throw new Exception(
                            $"No header found for column: {columnNo} on row: {rowNo}, cell address: {cell.Address}");

                    row.Add(header, Convert.ToString(cell.Value, cultureInfo));
                }

                result.Add(row);
            }

            return result;
        }
    }
}
