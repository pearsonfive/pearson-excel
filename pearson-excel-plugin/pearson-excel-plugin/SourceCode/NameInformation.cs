﻿using System.Diagnostics.Eventing.Reader;
using NetOffice.ExcelApi;

namespace Pearson.Excel.Plugin.SourceCode
{
    public class NameInformation
    {
        public Name Name { get; set; }
        public string RangeName => Name.Name;
        public string Address => Name.RefersTo as string;

        public string Scope
        {
            get
            {
                try
                {
                    return Name.Name.Contains(@"!") ? Name.RefersToRange.Worksheet.Name : "";
                }
                catch
                {
                    return "";
                }
            }
        }
    }
}