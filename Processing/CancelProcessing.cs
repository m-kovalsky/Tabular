#r "Microsoft.AnalysisServices.Core.dll"

var DMV_Cmd = ExecuteDax("SELECT [SESSION_ID],[SESSION_LAST_COMMAND] FROM $SYSTEM.DISCOVER_SESSIONS").Tables[0];
bool runTMSL = true;
string databaseID = Model.Database.ID;
string databaseName = Model.Database.Name;
string sID = string.Empty;

for (int r = 0; r < DMV_Cmd.Rows.Count; r++)
{
    string sessionID = DMV_Cmd.Rows[r][0].ToString();
    string cmdText = DMV_Cmd.Rows[r][1].ToString();
    
    // Capture refresh command for the database
    if (cmdText.StartsWith("<Batch Transaction=") && cmdText.Contains("<Refresh xmlns") && cmdText.Contains("<DatabaseID>"+databaseID+"</DatabaseID>"))
    {
        sID = sessionID;
    }      
}

if (sID == string.Empty)
{
    Error("No processing Session ID found for the '"+databaseName+"' model.");
    return;
}

if (runTMSL)
{
    Model.Database.TOMDatabase.Server.CancelSession(sID);
    Info("Processing for the '"+databaseName+"' model has been cancelled (Session ID: "+sID+").");
}
else
{
    sID.Output();
}