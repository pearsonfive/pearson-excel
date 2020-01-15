using System;
using ExcelDna.Integration.CustomUI;

namespace Pearson.Excel.Plugin.Ribbon
{
    public class DropDownRibbonControl: RibbonControl
    {
        public DropDownRibbonControl(IRibbonUI ribbon) : base("dropDown", ribbon)
        {
        }

        public Action<DropDownRibbonControl> OnAction { get; set; }
        public string SelectedItemId { get; set; }
        public int SelectedIndex { get; set; }


        //public interface ISelectable
        //{
        //    string SelectedItemId { get; set; }
        //    int SelectedIndex { get; set; }
        //}

    }
}