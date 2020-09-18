using CommandLine;

namespace Populate
{
    public class CommandLineOptions
    {
        [Option('t', "template", Required = true, HelpText = "Template file to populate with Excel file values.")]
        public string TemplateFile { get; set; }

        [Option('v', "values", Required = true, HelpText = "Excel file for values to populate template.")]
        public string ExcelFile { get; set; }

        [Option('o', "out", Required = true, HelpText = "Directory to save each file.")]
        public string OutputDirectory { get; set; }

        [Option('c', "filename-column", Required = false, HelpText = "Excel column for each file name.")]
        public string OutputFilenameColumn { get; set; }

        [Option('e', "extension", Required = false, Default = "txt", HelpText = "Output files extension.")]
        public string Extension { get; set; }
    }
}