using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using NetOffice.ExcelApi;
using Pearson.Excel.Plugin.Ribbon;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Pearson.Excel.Plugin
{
    [ComVisible(true)]
    public class RibbonHandler : ExcelRibbon
    {
        private static readonly Dictionary<string, RibbonControl> controls = new Dictionary<string, RibbonControl>();

        private static readonly Application app = new Application(null, ExcelDnaUtil.Application);
        private IRibbonUI _ribbon;

        public void Ribbon_Load(IRibbonUI sender)
        {
            _ribbon = sender;

            new List<RibbonControl>
            {
                new RibbonControl(_ribbon)
                {
                    Id = "CalculationGroup",
                    IsVisible = true
                },
                new RibbonControl(_ribbon)
                {
                    Id = "btnCalcNow",
                    Label = "Calculate Now",
                    ImageMso = "AcceptInvitation",
                    IsEnabled = true,
                    IsVisible = true
                },
                new RibbonControl(_ribbon)
                {
                    Id = "btnCalcSheet",
                    Label = "Calculate Sheet",
                    ImageMso = "AccessFormDatasheet",
                    IsEnabled = false,
                    IsVisible = true
                }
            }.ForEach(control => controls[control.Id] = control);
        }

        private void invalidateRibbon()
        {
            ExcelAsyncUtil.QueueAsMacro(() => _ribbon.Invalidate());
        }

        #region Get*****

        public string GetLabel(IRibbonControl control)
        {
            var c = controls[control.Id];
            return c.Label;
        }

        public string GetImage(IRibbonControl control)
        {
            var c = controls[control.Id];
            return c.ImageMso;
        }

        public bool GetVisible(IRibbonControl control)
        {
            var c = controls[control.Id];
            return c.IsVisible;
        }

        public bool GetEnabled(IRibbonControl control)
        {
            var c = controls[control.Id];
            return c.IsEnabled;
        }

        #endregion

        #region On***

        public void OnAction(IRibbonControl control)
        {

        }

#endregion
    }
}