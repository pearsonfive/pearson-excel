using System;
using ExcelDna.Integration;

namespace Pearson.Excel.Plugin
{
    public class Addin : IExcelAddIn
    {
        public void AutoOpen()
        {
            
        }

        public void AutoClose()
        {
            // this never gets called
            throw new NotImplementedException();
        }
    }
}
