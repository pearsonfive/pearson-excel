using ExcelDna.Integration.CustomUI;

namespace Pearson.Excel.Plugin.Ribbon
{
    public class RibbonControl
    {
        private readonly IRibbonUI _ribbon;
        private readonly string _type;
        public string Id { get; set; }

        public RibbonControl(string type, IRibbonUI ribbon)
        {
            _type = type;
            _ribbon = ribbon;
        }

        public RibbonControl(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }


    }
}