#r "Microsoft.AnalysisServices.Core.dll"
using ToM = Microsoft.AnalysisServices.Tabular;

string serverName = Model.Database.TOMDatabase.Server.ToString();
string cmdText = @"start /wait ""C:\Program Files (x86)\Tabular Editor"" TabularEditor.exe " + @"""" + serverName + @"""";
string folderPath = @"C:\Desktop\MyFolder"; // Folder where the .bim or save-to-folder files are saved
string saveType = "B"; // Use 'B' for saving to .bim files, use 'F' for saving to folder structure

foreach (var x in Model.Database.TOMDatabase.Server.Databases)
{
    string dbName = x.ToString();
    string fullCmdText = cmdText + @" """ + dbName + @""" -" + saveType + " " + @"""" + folderPath + @"\" + dbName + @".bim"""; 
    
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