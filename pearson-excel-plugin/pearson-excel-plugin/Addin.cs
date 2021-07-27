using System;
using ExcelDna.Integration;

namespace Pearson.Excel.Plugin
{
    public class Addin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex=>$"!!! EXCEPTION: {ex.ToString()}");

            var funcRegistration = new RemoteFunctions.FunctionRegistration();
            funcRegistration.Register();
        }

        public void AutoClose()
        {
            // this never gets called
            throw new NotImplementedException();
        }
    }
}
