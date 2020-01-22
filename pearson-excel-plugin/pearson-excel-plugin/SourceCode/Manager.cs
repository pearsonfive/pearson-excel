using System;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;
using Pearson.Excel.Plugin.Extensions;
using Application = NetOffice.ExcelApi.Application;

namespace Pearson.Excel.Plugin.SourceCode
{
    public class Manager
    {
        private static readonly Application app = new Application(null, ExcelDnaUtil.Application);
        private const string EXCEL_SOURCE_FOLDER = "source";

        public Manager() { }

        [ExcelCommand]
        public static void OutputNames(string outputPath)
        {
            var folder = createOutputFolder(outputPath);
            var workbook = app.ThisWorkbook;

            var names = workbook.Names;

            var header = new[]
            {
                $"{"Name",-55} {"Scope",-30} Address",
                $"{"====",-55} {"=====",-30} ======="
            };

            var namesList = names
                .Select(nm => new NameInformation {Name = nm})
                .OrderBy(nm => nm.RangeName)
                .Select(nm => $"{nm.RangeName,-55} {nm.Scope, -30} {nm.Address}");
            var result = header.Concat(namesList).Aggregate((temp, next) => temp + Environment.NewLine + next);

            var fullPath = Path.Combine(folder, "NamedRanges.txt");
            app.StatusBar = $"Writing file [{fullPath}]";
            File.WriteAllText(fullPath, result);
            app.StatusBar = false;
        }

        [ExcelCommand]
        public static void OutputCodeModules(string outputPath)
        {
            var folder = createOutputFolder(outputPath);
            var workbook = app.ThisWorkbook;

            // delete old files
            new[] {".bas", ".cls", ".frm", ".frx"}
                .ForEach(extension => Directory.GetFiles(folder, $"*{extension}").ForEach(File.Delete));

            var components = workbook.VBProject.VBComponents;
            components
                .Where(cpt => getFileExtension(cpt)=="frm" || countNonEmptyLines(cpt.CodeModule) > 0)
                .ForEach(cpt =>
                {
                    var filename = $"{cpt.Name}{getFileExtension(cpt)}";
                    app.StatusBar = $"Writing file [{Path.Combine(folder, filename)}]";
                    var code = cpt.CodeModule.Lines(1, cpt.CodeModule.CountOfLines).Trim().Trim('\n');
                    File.WriteAllText(Path.Combine(folder, filename), code);
                    app.StatusBar = false;
                });
        }

        private static string createOutputFolder(string path = null)
        {
            var folder = path ?? Path.Combine(app.ActiveWorkbook.Path, EXCEL_SOURCE_FOLDER);
            Directory.CreateDirectory(folder);
            return folder;
        }

        private static string getFileExtension(VBComponent component)
        {
            var extension = ".bas";
            switch (component.Type)
            {
                case vbext_ComponentType.vbext_ct_ClassModule:
                case vbext_ComponentType.vbext_ct_Document:
                {
                    extension = ".cls";
                    break;
                }
                case vbext_ComponentType.vbext_ct_MSForm:
                {
                    extension = ".frm";
                    break;
                }
                case vbext_ComponentType.vbext_ct_StdModule:
                {
                    extension = ".bas";
                    break;
                }
            }
            return extension;
        }

        private static int countNonEmptyLines(CodeModule codeModule)
        {
            var counter = 0;
            for (var k = 1; k <= codeModule.CountOfLines; k++)
            {
                var s = codeModule.Lines(k, 1);
                if (!string.IsNullOrEmpty(s)) counter++;
            }
            return counter;
        }
    }
}