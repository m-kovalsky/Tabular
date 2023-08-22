#r "Microsoft.AnalysisServices.Core.dll"

var DMV_Connections = ExecuteDax("SELECT [CONNECTION_ID],[CONNECTION_USER_NAME],[CONNECTION_LAST_COMMAND_ELAPSED_TIME_MS] FROM $SYSTEM.DISCOVER_CONNECTIONS").Tables[0];
var DMV_Sessions = ExecuteDax("SELECT [SESSION_SPID],[SESSION_CONNECTION_ID],[SESSION_USER_NAME] FROM $SYSTEM.DISCOVER_SESSIONS").Tables[0];
var DMV_Commands = ExecuteDax("SELECT [SESSION_SPID],[COMMAND_TEXT] FROM $SYSTEM.DISCOVER_COMMANDS").Tables[0];

int thresholdSec = 5*60; // Set the threshold (seconds)
string[] userNames = { "" }; // Enter an array of user names for which you do not want queries to be cancelled

thresholdSec = thresholdSec * 1000 // Convert to seconds

for (int r=0; r < DMV_Connections.Rows.Count; r++)
{
    string connID = DMV_Connections.Rows[r][0].ToString();
    string userName = DMV_Connections.Rows[r][1].ToString();
    int timeMS = Convert.ToInt32(DMV_Connections.Rows[r][2].ToString());
        
    if (!userNames.Contains(userName) && timeMS > thresholdSec) // do not cancel certain users
    {
        for (int a=0; a < DMV_Sessions.Rows.Count; a++)
        {
            string spid = DMV_Sessions.Rows[a][0].ToString();
            string sConnID = DMV_Sessions.Rows[a][1].ToString();
            
            for (int b=0; b < DMV_Commands.Rows.Count; b++)
            {
                string sp = DMV_Commands.Rows[b][0].ToString();
                string cmdText = DMV_Commands.Rows[b][1].ToString();
                
                if (connID == sConnID && sp == spid && !cmdText.StartsWith("<Batch Transaction=")) // do not cancel processing events
                {
                    string cmd = @"<Cancel xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine""><SPID>"+spid+"</SPID><CancelAssociated>true</CancelAssociated></Cancel>";
                                                        
                    try
                    {
                        ExecuteCommand(cmd,isXmla: true);
                        Info("SPID '" + spid + "' for user '" + userName + "' was cancelled.");
                    }
                    catch
                    {
                        Error("SPID '" + spid + "' was not found.");
                    }                                               
                }
            }                            
        }
    }
}