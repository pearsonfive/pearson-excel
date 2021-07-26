using System.Collections.Generic;


namespace Pearson.Excel.Plugin.RemoteFunctions
{
    public class MockPythonServer
    {
        public IEnumerable<FunctionInfo> GetFunctionInfos()
        {
            var infos = new List<FunctionInfo>
            {
                new FunctionInfo
                {
                    Description = "Description shown in the Function Wizard",
                    PythonFunctionName = "pyAddTwoNumbers",
                    NameForExcel = "autoregistered.addTwoNumbers",
                    DisableFunctionWizard = true,
                    IsAsync = false,
                    InputParams = new[]
                    {
                        new ArgInfo{Name="x", Type = "double"},
                        new ArgInfo{Name="y", Type = "double"}
                    }
                }
            };

            return infos;
        }
    }
}