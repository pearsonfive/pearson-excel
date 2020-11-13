public void RegisterFunctions()
{
    try
    {
        // Get all calculation details
        List<CalcInfo> calculations = Service.Instance.GetAllCalculations();
 
        // Create function registration entries using ExcelDNA.Registration
        IEnumerable<ExcelFunctionRegistration> calcEntries = calculations.Select(calc =>
        {
            // Create ExcelFunctionAttribute for function hints in excel
            ExcelFunctionAttribute funcAttr = new ExcelFunctionAttribute()
            {
                Name = calc.name,
                Description = calc.description
            };
 
            // Create parameter registration entries for parameter hint
            List<ExcelParameterRegistration> paramEntries = calc.inputParams.Select(p =>
                new ExcelParameterRegistration(new ExcelArgumentAttribute() { Name = p.name, Description = p.description })).ToList();
 
            // Create a lambda expression
            LambdaExpression exp = FuncToExpression(calc);
 
            // Return the registration instance
            return new ExcelFunctionRegistration(FuncToExpression(calc), funcAttr, paramEntries);
 
        });
 
        ExcelRegistration.RegisterFunctions(calcEntries);
    }
    catch (Exception ex)
    {
 
    }
}
         
/// <summary>
/// Converting our calculation into a LambdaExpression to be used by ExcelDNA.Registration
/// </summary>
/// <param name="calc">The calculation information</param>
/// <returns></returns>
private LambdaExpression FuncToExpression(CalcInfo calc)
{
    // Add as many cases as the maximum no. of arguments to be supported
    switch (calc.inputParams.Length)
    {
        case 1: return FuncExpression1((a1) => ExecuteCalculation(calc, a1));
        case 2: return FuncExpression2((a1, a2) => ExecuteCalculation(calc, a1, a2));
        case 3: return FuncExpression3((a1, a2, a3) => ExecuteCalculation(calc, a1, a2, a3));
        case 4: return FuncExpression4((a1, a2, a3, a4) => ExecuteCalculation(calc, a1, a2, a3, a4));
        case 5: return FuncExpression5((a1, a2, a3, a4, a5) => ExecuteCalculation(calc, a1, a2, a3, a4, a5));
        default:
            return null;
    }
}
 
// LambdaExpression generators for different no. of arguments
public Expression<Func<object, object>> FuncExpression1(Func<object, object> f) { return (a1) => f(a1); }
public Expression<Func<object, object, object>> FuncExpression2(Func<object, object, object> f) { return (a1, a2) => f(a1, a2); }
public Expression<Func<object, object, object, object>> FuncExpression3(Func<object, object, object, object> f) { return (a1, a2, a3) => f(a1, a2, a3); }
public Expression<Func<object, object, object, object, object>> FuncExpression4(Func<object, object, object, object, object> f) { return (a1, a2, a3, a4) => f(a1, a2, a3, a4); }
public Expression<Func<object, object, object, object, object, object>> FuncExpression5(Func<object, object, object, object, object, object> f) { return (a1, a2, a3, a4, a5) => f(a1, a2, a3, a4, a5); }
 
/// <summary>
/// Call the web service endpoint with proper request body
/// </summary>
/// <param name="calc">The information related to the calculation invoked</param>
/// <param name="args">Arguments passed by excel</param>
/// <returns></returns>
private object ExecuteCalculation(CalcInfo calc, params object[] args)
{
    // Your own implementation of executing any calculation by calling the service
    object result = null;
             
    return result;
}