using System;
using System.Collections.Generic;
using ExcelDna.Integration.CustomUI;

namespace Pearson.Excel.Plugin.Ribbon
{
    public class ToggleButtonRibbonControl : RibbonControl
    {
        public ToggleButtonRibbonControl(IRibbonUI ribbon) : base("toggleButton", ribbon)
        {
        }

        public bool IsPressed { get; set; }
        public Action<IRibbonControl, bool> OnToggle { get; set; }

        protected override List<Tuple<string, string>> GetAttributes()
        {
            var attributes = base.GetAttributes();
            attributes.Add(new Tuple<string, string>("onAction", "OnActionPressed"));
            attributes.Add(new Tuple<string, string>("getPressed", "GetPressedToggle"));
            return attributes;
        }
    }
}