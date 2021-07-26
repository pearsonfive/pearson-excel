

namespace Pearson.Excel.Plugin.RemoteFunctions
{
    public class FunctionInfo
    {
        public string PythonFunctionName { get; set; }
        public string NameForExcel { get; set; }
        public string Description { get; set; }
        public bool IsAsync { get; set; }
        public bool DisableFunctionWizard { get; set; }
        public ArgInfo[] InputParams { get; set; }
    }
}