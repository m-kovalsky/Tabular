// Initial parameters
string[] tableList = {"Calendar","Geography"}; // Enter tables to process. If you want to process the whole database, leave the array as {};
string type = "full"; // Options: 'full','automatic','calculate','clearValues'

// Additional parameters
string databaseName = Model.Database.Name;
string tmslStart = "{ \"refresh\": { \"type\": \""+type+"\",\"objects\": [ ";
string tmslMid = "{\"database\": \""+databaseName+"\",\"table\": \"%table%\"} ";
string tmslMidDB = "{\"database\": \""+databaseName+"\"}";
string tmslEnd = "] } }";
string tmsl = tmslStart;
string tablePrint = string.Empty;
string[] typeList = {"full","automatic","calculate","clearValues"};
var sw = new System.Diagnostics.Stopwatch();

// Error check: processing type
if (!typeList.Contains(type))
{
    Error("Invalid processing 'type'. Please enter a valid processing type within the 'type' parameter.");
    return;
}

// Error check: validate tables
for (int i=0; i<tableList.Length; i++)
{
    if (!Model.Tables.Any(a => a.Name == tableList[i]))
    {
        Error("'"+tableList[i]+"' is not a valid table in the model.");
        return;
    }
}

// Generate TMSL
if (tableList.Length == 0)
{
    tmsl = tmsl+tmslMidDB+tmslEnd;
}
else
{
    for (int i=0; i<tableList.Length; i++)
    {
        if (i == 0)
        {
            tmsl = tmsl+tmslMid.Replace("%table%",tableList[i]);
        }
        else
        {
            tmsl = tmsl+","+tmslMid.Replace("%table%",tableList[i]);
        }
        
        tablePrint = tablePrint + "'"+tableList[i]+"',";
    }
    
    tmsl = tmsl + tmslEnd;    
}

// Run TMSL and output info text
tablePrint = tablePrint.Substring(0,tablePrint.Length-1);
sw.Start();
ExecuteCommand(tmsl);
sw.Stop();
Info("Processing '"+type+"' of tables: ["+tablePrint+"] finished in: " + sw.ElapsedMilliseconds + " ms");
