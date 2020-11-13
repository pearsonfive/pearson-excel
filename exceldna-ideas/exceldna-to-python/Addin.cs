public void AutoOpen()
{
    try
    {
        IntelliSenseServer.Install();
        RegisterFunctions();
        // Refresh Excel Intellisense
        IntelliSenseServer.Refresh();
        xlApp.StatusBar = "Functions registered";
    }
    catch (Exception ex)
    {
        xlApp.StatusBar = "Error in registering functions";
    }
}
 
public void AutoClose()
{
    IntelliSenseServer.Uninstall();
}
 
public void RegisterFunctions()
{
}