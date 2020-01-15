using System;
using System.Collections.Generic;
using ExcelDna.Integration.CustomUI;

namespace Pearson.Excel.Plugin.Ribbon
{
    public class ButtonRibbonControl : RibbonControl
    {
        public ButtonRibbonControl(IRibbonUI ribbon) : base("button", ribbon)
        {
        }

        protected override List<Tuple<string, string>> GetAttributes()
        {
            var attributes = base.GetAttributes();
            attributes.Add(new Tuple<string, string>("onAction", "OnAction"));
            attributes.Add(new Tuple<string, string>("getImage", "GetImage"));
            return attributes;
        }

        public Action<IRibbonControl> OnAction { get; set; }
    }
}