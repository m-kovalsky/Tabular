#r "System.IO"
#r "System.IO.Compression.FileSystem"

using System.IO;
using System.IO.Compression;

bool createMeasures = false; // Set this to true if you want the measures to be created within the model. If set to false it will generate a script which can be executed to create the measures.
string pbiFile = @"C:\Desktop\ReportFile.pbix"; // Update this to a .pbix or .pbit file.
string fileExt = Path.GetExtension(pbiFile);
string fileName = Path.GetFileNameWithoutExtension(pbiFile);
string folderName = Path.GetDirectoryName(pbiFile) + @"\";
string zipPath = folderName + fileName + ".zip";
string unzipPath = folderName + fileName;

if (! (fileExt == ".pbix" || fileExt == ".pbit"))
{
   Error("Must enter a valid .pbix or .pbit file");
   return;
}

try
{
   // Make a copy of a pbi and turn it into a zip file
   File.Copy(pbiFile, zipPath);
   // Unzip file
   System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, unzipPath);
   // Delete zip file
   File.Delete(zipPath);
}
catch
{
   Error("File does not exist. Must use a valid .pbix or .pbit file");
   return;
}

string layoutPath = unzipPath + @"\Report\Layout";
string jsonFilePath = Path.ChangeExtension(layoutPath, ".json");
File.Move(layoutPath, jsonFilePath); 

string unformattedJson = File.ReadAllText(jsonFilePath,System.Text.UnicodeEncoding.Unicode);
string formattedJson = Newtonsoft.Json.Linq.JToken.Parse(unformattedJson).ToString();
dynamic json = Newtonsoft.Json.Linq.JObject.Parse(formattedJson);

// Delete previously created folder
try
{
    Directory.Delete(folderName + fileName,true);
}
catch
{
}

var sb = new System.Text.StringBuilder();
var sb_Script = new System.Text.StringBuilder();
string newline = Environment.NewLine;
sb.Append("TableName" + '\t' + "MeasureName" + '\t' + "Expression" + '\t' + "HiddenFlag" + '\t' + "FormatString" + '\t' + "MeasureCreatedFlag" + newline);

string config = (string)json["config"];
string formattedconfigJson = Newtonsoft.Json.Linq.JToken.Parse(config).ToString();
dynamic configJson = Newtonsoft.Json.Linq.JObject.Parse(formattedconfigJson);

int i=0;

try{
    foreach (var o in configJson["modelExtensions"].Children())
    {    
        foreach (var o2 in o["entities"].Children())
        {
            string tableName = (string)o2["name"];
            
            foreach (var o3 in o2["measures"].Children())
            {
                string measureName = (string)o3["name"];
                string expr = (string)o3["expression"];
                bool hid = (bool)o3["hidden"];
                string fs = (string)o3["formatInformation"]["formatString"];
                
                sb.Append(tableName + '\t' + measureName + '\t' + expr + '\t' + hid + '\t' + fs + '\t');
                
                if (createMeasures)
                {
                    try
                    {
                        if (Model.AllMeasures.Any(a => a.Name == measureName))
                        {
                            Warning("Unable to create the '"+measureName+"' measure as this measure already exists in the model.");
                            sb.Append("No" + newline);
                        }
                        else
                        {
                            var m = Model.Tables[tableName].AddMeasure(measureName);
                            m.Expression = expr;
                            m.IsHidden = hid;
                            m.FormatString = fs;
                            
                            sb.Append("Yes" + newline);
                        }
                    }
                    catch
                    {
                        Warning("Unable to create the '"+measureName+"' measure as the '"+tableName+"' table does not exist.");
                        sb.Append("No" + newline);
                    }                
                }
                else
                {
                    expr = expr.Replace("\"",@"\""").Replace("\r","\\r").Replace("\n","\\n").Replace("\t","\\t");
                    
                    sb_Script.Append("var m"+i+" = Model.Tables[\"" + tableName + "\"].AddMeasure(\"" + measureName + "\");" + newline);
                    sb_Script.Append("m"+i+".Expression = \"" + expr + "\";" + newline);
                    sb_Script.Append("m"+i+".IsHidden = " + hid.ToString().ToLower() + ";" + newline);
                    sb_Script.Append("m"+i+".FormatString = @\"" + fs + "\";" + newline + newline);
                }
                i++;
            }        
        }    
    }
}
catch
{
    Error("There are no report level measures in this report. All the measures in this report are in the model itself.");
}

sb.Output(); // Outputs a list of the report-level measures

if (!createMeasures)
{
    sb_Script.Output(); // Outputs a C# script which can be executed in the Advanced Scripting window to create the measures
}

