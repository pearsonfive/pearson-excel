using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using NetOffice.ExcelApi;
using Pearson.Excel.Plugin.Ribbon;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Pearson.Excel.Plugin
{
    [ComVisible(true)]
    public class RibbonHandler: ExcelRibbon
    {
        private static readonly Dictionary<string, RibbonControl> controls = new Dictionary<string, RibbonControl>();

        private static readonly  Application app = new Application(null, ExcelDnaUtil.Application);
        private IRibbonUI _ribbon;

        public void Ribbon_Load(IRibbonUI sender)
        {
            _ribbon = sender;

            new List<RibbonControl>
            {
                new RibbonControl(_ribbon)
                {
                    Id = "btnCalcNow"
                },
                new RibbonControl(_ribbon)
                {
                    Id = "btnCalcSheet"
                }
            }.ForEach(control => controls[control.Id] = control);
        }
    }
}