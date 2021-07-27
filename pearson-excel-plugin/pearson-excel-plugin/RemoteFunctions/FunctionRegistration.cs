using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reactive.Linq;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Registration;
using ExcelDna.Registration.Utils;
using NetOffice.ExcelApi;

namespace Pearson.Excel.Plugin.RemoteFunctions
{
    public class FunctionRegistration
    {
        /// <summary>
        /// Registers the python functions with ExcelDna
        /// </summary>
        public void Register()
        {
            try
            {
                var server = new MockPythonServer();
                var funcInfos = server.GetFunctionInfos();

                var funcEntries = funcInfos.Select(info =>
                {
                    var attr = new ExcelFunctionAttribute
                    {
                        Name = info.NameForExcel,
                        Description = info.Description
                    };

                    var parms = info.InputParams.Select(p =>
                        new ExcelParameterRegistration(new ExcelArgumentAttribute
                            { Name = p.Name, Description = p.Description }));

                    var expr = info.IsAsync ? funcToExpressionAsync(info) : funcToExpression(info);

                    return new ExcelFunctionRegistration(expr, attr, parms);
                });

                var postAsyncReturnConfig = GetPostAsyncReturnConversionConfig();

                funcEntries
                    .ProcessParameterConversions(postAsyncReturnConfig)
                    .RegisterFunctions();
            }
            catch (Exception)
            {

            }
        }

        static ParameterConversionConfiguration GetPostAsyncReturnConversionConfig()
        {
            // This conversion replaces the default #N/A return value of async functions with the #GETTING_DATA value.
            // This is not supported on old Excel versions, bu looks nicer these days.
            // Note that this ReturnConversion does not actually check whether the functions is an async function, 
            // so all registered functions are affected by this processing.
            return new ParameterConversionConfiguration()
                .AddReturnConversion((type, customAttributes) => type != typeof(object)
                    ? null
                    : (Expression<Func<object, object>>) (returnValue => returnValue.Equals(ExcelError.ExcelErrorNA)
                        ? "#BUSY" // or any other value, e.g. "#calculating..." or ExcelError.ExcelErrorGettingData
                        : returnValue));
        }

        private LambdaExpression funcToExpression(FunctionInfo info)
        {
            switch (info.InputParams.Length)
            {
                case 1: return FuncExpression1(a1 => executeFunction(info, a1));
                case 2: return FuncExpression2((a1, a2) => executeFunction(info, a1, a2));
                // etc
                // etc
                // etc
                default: return null;
            }
        }

        private LambdaExpression funcToExpressionAsync(FunctionInfo info)
        {
            switch (info.InputParams.Length)
            {
                case 1: return FuncExpression1(a1 => executeFunctionAsync(info, a1));
                case 2: return FuncExpression2((a1, a2) => executeFunctionAsync(info, a1, a2));
                // etc
                // etc
                // etc
                default: return null;
            }
        }

        public Expression<Func<object, object>> FuncExpression1(Func<object, object> f)
        {
            return a1 => f(a1);
        }
        public Expression<Func<object, object, object>> FuncExpression2(Func<object, object, object> f)
        {
            return (a1, a2) => f(a1, a2);
        }
        // etc
        // etc
        // etc

        private object executeFunction(FunctionInfo info, params object[] args)
        {
            if (info.DisableFunctionWizard && ExcelDnaUtil.IsInFunctionWizard())
                return "### Not while in Function Wizard ###";

            object result = $"{info.NameForExcel}(): do some calcs";
            //object result = callApi(info.PythonFunctionName, args);

            return result;
        }

        private object executeFunctionAsync(FunctionInfo info, params object[] args)
        {
            return ExcelAsyncUtil.Run(info.NameForExcel, args, delegate
            {
                Thread.Sleep(1000);
                return "Function complete";
            });

            //return ObservableRtdUtil.Observe(info.NameForExcel, args, GetObservableClock);

        }

        //static IObservable<string> GetObservableClock()
        //{
        //    return Observable.Timer(dueTime: TimeSpan.Zero, period: TimeSpan.FromSeconds(1))
        //        .Select(_ => DateTime.Now.ToString("HH:mm:ss"));
        //}

    }
}