#r "Microsoft.AnalysisServices.Core.dll"

// Initial Parameters
bool runTMSL = true;
string processingType = "full"; // Enter the processing type
string[] typeList = {"full","automatic","calculate","clearValues","defragment","dataOnly"};
string[] tableOmit = {}; // Set tables to omit from the processing script
bool seqEnabled = false; // Indicate whether to use the sequence command
int maxP = 0; // Enter the Max Parallelism (applicable if using sequence command)
string databaseName = Model.Database.Name;
var DMV_Tables = ExecuteDax("SELECT [ID],[Name] FROM $SYSTEM.TMSCHEMA_TABLES").Tables[0];
var DMV_Partitions = ExecuteDax("SELECT [ID],[TableID],[Name] FROM $SYSTEM.TMSCHEMA_PARTITIONS WHERE [State] <> 1").Tables[0];
string processingMethod = string.Empty; // Database, Table, Partition
string newline = Environment.NewLine;
TOM.SaveOptions so = new TOM.SaveOptions();
var sw = new System.Diagnostics.Stopwatch();
string timeSpent = "";

// Remove existing annotations
foreach (var o in Model.Tables.ToList())
{
    o.RemoveAnnotation("Vertipaq_TableID");
}

foreach (var o in Model.AllPartitions)
{
    o.RemoveAnnotation("Vertipaq_PartitionID");
}

// Add TableId annotations
for (int r = 0; r < DMV_Tables.Rows.Count; r++)
{
    string tblID = DMV_Tables.Rows[r][0].ToString();
    string tblName = DMV_Tables.Rows[r][1].ToString();
    
    Model.Tables[tblName].SetAnnotation("Vertipaq_TableID",tblID);         
}

// Add PartitionId annotations
for (int r = 0; r < DMV_Partitions.Rows.Count; r++)
{
    string pID = DMV_Partitions.Rows[r][0].ToString();
    string tableID = DMV_Partitions.Rows[r][1].ToString();
    string pName = DMV_Partitions.Rows[r][2].ToString();
    
    foreach (var t in Model.Tables.Where(a => a.GetAnnotation("Vertipaq_TableID") == tableID))
    {
        string tableName = t.Name;
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_PartitionID",pID);        
    }
}

// Error catching: Max Parallelism
if (seqEnabled == true && maxP <1)
{
    Error("Must enter a valid Max Parallelism value in the maxP parameter.");
    return;
}

// Update annotations based on tables/partitions to process
foreach (var t in Model.Tables.ToList())
{
    if (t.Partitions.All(a => a.HasAnnotation("Vertipaq_PartitionID")))
    {
        foreach (var p in t.Partitions.ToList())
        {
            p.RemoveAnnotation("Vertipaq_PartitionID");
        }
    }
    else
    {
        t.RemoveAnnotation("Vertipaq_TableID");
    }
}

// Identify processing method
if (Model.Tables.All(a => a.HasAnnotation("Vertipaq_TableID")))
{
    processingMethod = "Database";
}
else if (Model.AllPartitions.Any(a => a.HasAnnotation("Vertipaq_PartitionID")))
{
    processingMethod = "Partition";
}
else if (Model.Tables.Any(a => a.HasAnnotation("Vertipaq_TableID")))
{
    processingMethod = "Table";
}
else
{
    Info("All tables/partitions are already processed.");
    return;
}

// Error check: processing type
if (!typeList.Contains(processingType))
{
    Error("Invalid processing 'type'. Please enter a valid processing type within the 'type' parameter.");
    return;
}

// Determine refresh type
var refType = TOM.RefreshType.Full;

if (processingType.ToLower() == "automatic")
{
    refType = TOM.RefreshType.Automatic;
}
else if (processingType.ToLower() == "dataonly")
{
    refType = TOM.RefreshType.DataOnly;
}
else if (processingType.ToLower() == "clearvalues")
{
    refType = TOM.RefreshType.ClearValues;
}
else if (processingType.ToLower() == "calculate")
{
    refType = TOM.RefreshType.Calculate;
}
else if (processingType.ToLower() == "defragment")
{
    refType = TOM.RefreshType.Defragment;
}

// Build Info output text
var sb_Info = new System.Text.StringBuilder();
sb_Info.Append("Processing type '"+processingType+"' of the '"+databaseName+"' model ");

// Generate request refresh
if (processingMethod == "Database")
{
    Model.Database.TOMDatabase.Model.RequestRefresh(refType);
}
else if (processingMethod == "Table")
{
    sb_Info.Append("for the following tables: [");
    foreach (var t in Model.Tables.Where(a => a.HasAnnotation("Vertipaq_TableID")))
    {
        string tableName = t.Name;
        Model.Database.TOMDatabase.Model.Tables[tableName].RequestRefresh(refType);   
        sb_Info.Append("'"+tableName+"',");
    }

    sb_Info.Remove(sb_Info.Length-1,1);
    sb_Info.Append("]");
}
else if (processingMethod == "Partition")
{
    sb_Info.Append("for the following partitions: [");    
    foreach (var t in Model.Tables.Where(a => a.HasAnnotation("Vertipaq_TableID") || a.Partitions.Any(b => b.HasAnnotation("Vertipaq_PartitionID"))))
    {
        string tableName = t.Name;
        
        if (t.HasAnnotation("Vertipaq_TableID"))
        {
            foreach (var p in t.Partitions.ToList())
            {
                string pName = p.Name;
                Model.Database.TOMDatabase.Model.Tables[tableName].Partitions[pName].RequestRefresh(refType);
                sb_Info.Append("'"+tableName+"'["+pName+"],");
            }
        }
        else
        {
            foreach (var p in t.Partitions.Where(a => a.HasAnnotation("Vertipaq_PartitionID")))
            {
                string pName = p.Name;
                Model.Database.TOMDatabase.Model.Tables[tableName].Partitions[pName].RequestRefresh(refType);
                sb_Info.Append("'"+tableName+"'["+pName+"],");             
            }
        }
    }

    sb_Info.Remove(sb_Info.Length-1,1);
    sb_Info.Append("]");
}

sb_Info.Append(" has finished in ");

// Remove Annotations
foreach (var o in Model.Tables.ToList())
{
    o.RemoveAnnotation("Vertipaq_TableID");
}

foreach (var o in Model.AllPartitions)
{
    o.RemoveAnnotation("Vertipaq_PartitionID");
}

if (runTMSL)
{
    sw.Start();
    // Add sequence if it is enabled
    if (seqEnabled)
    {        
        so.MaxParallelism = maxP;        
    }
    Model.Database.TOMDatabase.Model.SaveChanges(so); 
    sw.Stop();

    TimeSpan ts = sw.Elapsed;
  
    int sec = ts.Seconds;
    int min = ts.Minutes;
    int hr = ts.Hours;

    // Break down hours,minutes,seconds
    if (hr == 0)
    {
        if (min == 0)
        {
            timeSpent = sec + " seconds.";
        }
        else
        {
            timeSpent = min + " minutes and " + sec + " seconds.";
        }
    }
    else
    {
        timeSpent = hr + " hours, " + min + " minutes and " + sec + " seconds.";
    }

    if (hr == 1)
    {
        timeSpent = timeSpent.Replace("hours","hour");
    }
    if (min == 1)
    {
        timeSpent = timeSpent.Replace("minutes","minute");
    }
    if (sec == 1)
    {
        timeSpent = timeSpent.Replace("seconds","second");
    }

    Info(sb_Info.ToString() + timeSpent);
    return;
}
else
{
    if (processingMethod == "Database")
    {
        var x = Model.Database.TOMDatabase;
        TOM.JsonScripter.ScriptRefresh(x,refType).Output();
    }
    else if (processingMethod == "Table")
    {
        var x = Model.Database.TOMDatabase.Model.Tables.Where(a => a.Annotations.Where(b => b.Name == "Vertipaq_TableID").Count() == 1).ToArray();
        TOM.JsonScripter.ScriptRefresh(x,refType).Output();
    }
    else if (processingMethod == "Partition")
    {
    }
}