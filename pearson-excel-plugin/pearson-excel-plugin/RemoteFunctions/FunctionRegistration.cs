using System;
using System.Linq;
using System.Linq.Expressions;
using ExcelDna.Integration;
using ExcelDna.Registration;
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

                    var expr = funcToExpression(info);

                    return new ExcelFunctionRegistration(expr, attr, parms);
                });

                funcEntries.RegisterFunctions();
            }
            catch (Exception)
            {

            }
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


    }
}