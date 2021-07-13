using System;
using NetOffice.ExcelApi;

namespace Pearson.Excel.Plugin.NameManager
{
    public class NameInformation
    {
        public Name SourceName
        {
            set
            {
                Name = value.Name;
                refersToLocal = value.RefersToLocal as string ;
                refersToRange = value.RefersToRange;
            }
        }

        public Worksheet Worksheet
        {
            get
            {
                if (!(IsRefError || IsFormula)) return refersToRange.Worksheet;
                return null;
            }
        }
        public Workbook Workbook
        {
            get
            {
                if (!(IsRefError || IsFormula)) return Worksheet.Parent as Workbook;
                return null;
            }
        }
        public Range Range
        {
            get
            {
                if (!(IsRefError || IsFormula)) return refersToRange;
                return null;
            }
        }

        public string ShortAddress
        {
            get
            {
                if (!(IsRefError || IsFormula))
                    return LongAddress.Replace(WorksheetName, "").Replace("'", "").Replace("!", "");
                return LongAddress;
            }
        }

        public string Name { get; set; }
        public string LongName => Name;
        public string ShortName => bang > 0 ? LongName.Substring(bang + 1) : LongName;
        public string WorksheetName => Worksheet?.Name;
        public string WorkbookName => Workbook?.Name;
        public string Scope => Name.Contains(@"!") ? "Local" : "Global";
        public bool IsRefError => refersToLocal.Contains(@"#REF!");
        public bool IsFormula => refersToRange == null && !IsRefError;
        public string LongAddress => refersToLocal.Replace("$", "").Replace("=", "");


        private string refersToLocal { get; set; }
        private Range refersToRange { get; set; }
        private int bang => LongName.IndexOf(@"!", StringComparison.Ordinal);

    }
}