using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace Pearson.Excel.Plugin.Ribbon
{
    public class RibbonControl
    {
        private readonly IRibbonUI _ribbon;
        private readonly string _type;
        public string Id { get; set; }
        public string Label { get; set; }
        public string ImageMso { get; set; }
        public bool IsEnabled { get; set; }
        public bool IsVisible { get; set; }
        public Action OnInvalidate { get; set; }

        public RibbonControl(string type, IRibbonUI ribbon)
        {
            _type = type;
            _ribbon = ribbon;
        }

        public RibbonControl(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public void Invalidate()
        {
            OnInvalidate?.Invoke();
            ExcelAsyncUtil.QueueAsMacro(() => _ribbon.InvalidateControl(Id));
        }
    }
}