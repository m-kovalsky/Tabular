#r "Microsoft.AnalysisServices.Core.dll"
using ToM = Microsoft.AnalysisServices.Tabular;

string folderPath = @"C:\Desktop\MyFolder"; // Folder where the .bim or save-to-folder files are saved
string saveType = "B"; // Use 'B' for saving to .bim files, use 'F' for saving to folder structure

// If saving datasets from Power BI Premium, enter in your Service Principal credentials in the 3 parameters below:
string appID = "";
string tenantID = "";
string appSecret = "";

string serverName = Model.Database.TOMDatabase.Server.ToString();
string cmdText = @"start /wait /d ""C:\Program Files (x86)\Tabular Editor"" TabularEditor.exe " + @"""" + serverName + @"""";

// Update cmdText for Power BI Premium datasets
if (Model.DefaultPowerBIDataSourceVersion == PowerBIDataSourceVersion.PowerBI_V3)
{
    cmdText = @"start /wait /d ""C:\Program Files (x86)\Tabular Editor"" TabularEditor.exe " + @"""" + @"Provider=MSOLAP;Data Source=powerbi://api.powerbi.com/v1.0/myorg/" + serverName + ";User ID=app:" + appID + "@" + tenantID + ";Password=" + appSecret + @"""";
}

foreach (var x in Model.Database.TOMDatabase.Server.Databases)
{
    string dbName = x.ToString();
    string fullCmdText = cmdText + @" """ + dbName + @""" -" + saveType + " " + @"""" + folderPath + @"\" + dbName;
    
    
    if (saveType == "B")
    {
        fullCmdText = fullCmdText + @".bim""";
    }
    
    System.Diagnostics.Process process = new System.Diagnostics.Process();
    System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo("cmd");
    startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
    startInfo.FileName = "cmd.exe";
    startInfo.Arguments = fullCmdText;
    process.StartInfo = startInfo;    
    process.StartInfo.CreateNoWindow = true;
    process.StartInfo.RedirectStandardInput = true;
    process.StartInfo.UseShellExecute = false;
    process.Start();
    process.StandardInput.WriteLine(fullCmdText);
    process.StandardInput.Flush();
    process.StandardInput.Close();
    process.WaitForExit();
}