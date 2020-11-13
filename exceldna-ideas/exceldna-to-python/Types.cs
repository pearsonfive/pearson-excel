public class CalcInfo
{
    public string name { get; set; }
    public string description { get; set; }
    public ArgInfo[] inputParams { get; set; }
}
 
public class ArgInfo
{
    public string name { get; set; }
    public string type { get; set; }
    public string description { get; set; }
}