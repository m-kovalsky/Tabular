#r "System.IO"

using System.IO;
using TabularEditor.BestPracticeAnalyzer;

string folderPath = @"C:\Desktop\MyReport"; // Enter the folder with your model files
string textFilePath = folderPath + @"\BPAResults.txt"; // This is where the output .txt file is saved

var sb = new System.Text.StringBuilder();
string newline = Environment.NewLine;
sb.Append("ModelName" + '\t' + "RuleCategory" + '\t' + "RuleName" + '\t' + "ObjectName" + '\t' + "ObjectType" + '\t' + "RuleSeverity" + '\t' + "HasFixExpression" + '\t' + "RuleID" + newline);

// Extract all model files from the folderPath
string fileTypeBIM = "*.bim";
string fileTypeDJSON = "database.json";
string fileTypeTMDL = "model.tmd";

if (Directory.Exists(folderPath) == false)
{
    Error("Folder path specified in the 'folderPath' parameter is not valid. Please enter a valid folder path.");
    return;
}

string[] files = Directory.GetFiles(folderPath, fileTypeBIM, SearchOption.AllDirectories).Concat(Directory.GetFiles(folderPath, fileTypeDJSON, SearchOption.AllDirectories)).Concat(Directory.GetFiles(folderPath, fileTypeTMDL, SearchOption.AllDirectories)).ToArray();

if (files.Count() == 0)
{
    Warning("No model files found in the specified folder. Please specify a folder with one or more model files.");
    return;
}

// Loop through each model file and run BPA against the file
foreach (string filePath in files)
{
    // Extract model name
    string fileFolder = Path.GetFileName(Path.GetDirectoryName(filePath));
    string modelName = "";
    
    if (filePath.EndsWith(fileTypeDJSON) || filePath.EndsWith(fileTypeTMDL))
    {
        modelName = fileFolder;
    }
    else if (fileFolder.Contains(".Dataset"))
    {
        modelName = fileFolder.Substring(0,fileFolder.IndexOf(".Dataset"));
    }
    else if (Path.GetDirectoryName(filePath) == folderPath)
    {
        modelName = Path.GetFileNameWithoutExtension(filePath);
    }
    else
    {
        modelName = fileFolder;
    }
    
    // Run BPA against file
    var bpa = new Analyzer();
    var model = new TabularModelHandler(filePath).Model;    
    bpa.SetModel(model);
    
    foreach (var a in bpa.AnalyzeAll().ToList())
    {
        sb.Append(modelName + '\t' + a.Rule.Category + '\t' + a.RuleName + '\t' + a.ObjectName + '\t' + a.ObjectType + '\t' + a.Rule.Severity + '\t' + a.CanFix + '\t' + a.Rule.ID + newline);
    }
}

System.IO.File.WriteAllText(textFilePath, sb.ToString());

Info("BPA Scan Completed. File saved in this folder: " + folderPath);