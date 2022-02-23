using System;
using System.Reactive.Linq;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace Pearson.Excel.Plugin
{
    public class Addin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex=>$"!!! EXCEPTION: {ex.ToString()}");

            ExcelRegistration.GetExcelFunctions()
                .ProcessAsyncRegistrations()
                .RegisterFunctions();
        }

        [ExcelAsyncFunction]
        public static IObservable<object> getInfiniteStream(double sleepInterval, object dummy)
        {
            var counter = 0;

            return Observable.Interval(TimeSpan.FromSeconds(sleepInterval))
                .Select(_ =>
                {
                    counter++;
                    return counter as object;
                });
        }

        public void AutoClose()
        {
            throw new NotImplementedException();
        }

    }
}
