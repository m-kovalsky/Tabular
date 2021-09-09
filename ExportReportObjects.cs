#r "System.IO"
#r "System.IO.Compression.FileSystem"

using System.IO;
using System.IO.Compression;

/************************************************************************************************************

    Instructions:

    1. Paste this code into the Advanced Scripting window of Tabular Editor.
    2. Create a new model in Tabular Editor (File -> New Model) (Compatibility Level does not matter).
    3. Choose a method below and follow its respective steps.

    Method 1: For a single Power BI Desktop (.pbix) or Power BI Template (.pbit) file:
        1. Enter the folder location of the file in the 'pbiFolderName' parameter.
        2. Enter the file name (including the extension) in the 'pbiFile' parameter.

    Method 2: For looping through multiple .pbix or .pbit files within a folder:
        1. Enter the location of the folder in the 'pbiFolderName' parameter.
        2. Ensure that the pbiFile parameter is empty (should be this:  @""; ).

    Setting the 'saveToFile' parameter to 'true' will save the output text files to the respective folder.
    Setting the 'saveToFile' parameter to 'false' will generate the output within pop-up windows. 

************************************************************************************************************/

// User Parameters
string pbiFolderName = @"C:\Desktop\MyReports";
string pbiFile = @"MyReport.pbix";
bool saveToFile = true;
string savePrefix = "ReportObjects";

string newline = Environment.NewLine;
List<string> FileList = new List<string>();

var sb_CustomVisuals = new System.Text.StringBuilder();
sb_CustomVisuals.Append("ReportName" + '\t' + "CustomVisualName" + newline);

var sb_ReportFilters = new System.Text.StringBuilder();
sb_ReportFilters.Append("ReportName" + '\t' + "FilterName" + '\t' + "TableName" + '\t' + "ObjectName" + '\t' + "ObjectType" + '\t' + "FilterType" + newline);

var sb_VisualObjects = new System.Text.StringBuilder();
sb_VisualObjects.Append("ReportName" + '\t' + "PageName" + '\t' + "VisualId" + '\t' + "VisualType" + '\t' + "CustomVisualFlag" + '\t' + "TableName" + '\t' + "ObjectName" + '\t' + "ObjectType" + newline);

var sb_VisualFilters = new System.Text.StringBuilder();
sb_VisualFilters.Append("ReportName" + '\t' + "PageName" + '\t' + "VisualId" + '\t' + "TableName" + '\t' + "ObjectName" + '\t' + "ObjectType" + '\t' + "FilterType" + newline);

var sb_PageFilters = new System.Text.StringBuilder();
sb_PageFilters.Append("ReportName" + '\t' + "PageId" + '\t' + "PageName" + '\t' + "FilterName" + '\t' + "TableName" + '\t' + "ObjectName" + '\t' + "ObjectType" + '\t' + "FilterType" + newline);

var sb_Bookmarks = new System.Text.StringBuilder();
sb_Bookmarks.Append("ReportName" + '\t' + "BookmarkName" + '\t' + "BookmarkId" + '\t' + "PageId" + newline);

var sb_Pages = new System.Text.StringBuilder();
sb_Pages.Append("ReportName" + '\t' + "PageId" + '\t' + "PageName" + '\t' + "PageNumber" + '\t' + "PageWidth" + '\t' + "PageHeight" + '\t' + "PageHiddenFlag" + '\t' + "VisualCount" + newline);

var sb_Visuals = new System.Text.StringBuilder();
sb_Visuals.Append("ReportName" + '\t' + "PageName" + '\t' + "VisualId" + '\t' + "VisualName" + '\t' + "VisualType" + '\t' + "CustomVisualFlag" + '\t' + "VisualHiddenFlag" + '\t' + "X_Coordinate" + '\t' + "Y_Coordinate" + '\t' + "Z_Coordinate" + '\t' + "VisualWidth" + '\t' + "VisualHeight" + '\t' + "ObjectCount" + newline);

var sb_Connections = new System.Text.StringBuilder();
sb_Connections.Append("ReportName" + '\t' + "ServerName" + '\t' + "DatabaseName" + '\t' + "ConnectionType" + newline);

if (pbiFile.Length > 0 && pbiFolderName.Length == 0)
{
    Error("If specifying the 'pbiFile' you must also specify the 'pbiFolderName'.");
    return;
}    
else if (pbiFile.Length == 0 && pbiFolderName.Length > 0)
{
    foreach (var x in System.IO.Directory.GetFiles(pbiFolderName, "*.pbi"))
    {
        FileList.Add(x);
    }
}
else
{
    FileList.Add(pbiFolderName + @"\" + pbiFile);
}

foreach (var rpt in FileList)
{
    var CustomVisuals = new List<CustomVisual>();
    var Bookmarks = new List<Bookmark>();
    var ReportFilters = new List<ReportFilter>();
    var Visuals = new List<Visual>();
    var VisualObjects = new List<VisualObject>();
    var VisualFilters = new List<VisualFilter>();
    var PageFilters = new List<PageFilter>();
    var Pages = new List<Page>();
    var Connections = new List<Connection>();
    string fileExt = Path.GetExtension(rpt);
    string fileName = Path.GetFileNameWithoutExtension(rpt);
    string folderName = Path.GetDirectoryName(rpt) + @"\";
    string zipPath = folderName + fileName + ".zip";
    string unzipPath = folderName + fileName;

    if (! (fileExt == ".pbix" || fileExt == ".pbit"))
    {
        Error("'" +rpt+ "is not a valid file. File(s) must be a valid .pbix or .pbit.");
        return;
    }

    try
    {
        // Make a copy of a pbi and turn it into a zip file
        File.Copy(rpt, zipPath);
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

    // Layout file
    string layoutPath = unzipPath + @"\Report\Layout";
    string jsonFilePath = Path.ChangeExtension(layoutPath, ".json");
    File.Move(layoutPath, jsonFilePath); 

    string unformattedJson = File.ReadAllText(jsonFilePath,System.Text.UnicodeEncoding.Unicode);
    string formattedJson = Newtonsoft.Json.Linq.JToken.Parse(unformattedJson).ToString();
    dynamic json = Newtonsoft.Json.Linq.JObject.Parse(formattedJson);

    // Connections file
    string connPath = unzipPath + @"\Connections";
    string jsonconnFilePath = Path.ChangeExtension(connPath, ".json");
    File.Move(connPath, jsonconnFilePath); 

    string unformattedconnJson = File.ReadAllText(jsonconnFilePath,System.Text.Encoding.UTF8);
    string formattedconnJson = Newtonsoft.Json.Linq.JToken.Parse(unformattedconnJson).ToString();
    dynamic connjson = Newtonsoft.Json.Linq.JObject.Parse(formattedconnJson);

    //Delete previously created folder
    try
    {
        Directory.Delete(folderName + fileName,true);
    }
    catch
    {
    }

    string svName = string.Empty;
    string dbName = string.Empty;
    string connType = string.Empty;
    
    // Connection info
    try
    {
        foreach (var o in connjson["Connections"].Children())
        {
            connType = (string)o["ConnectionType"];
            try
            {
                
                dbName = (string)o["PbiModelDatabaseName"];
            }
            catch
            {
            }
            if (connType != "pbiServiceLive")
            {
                try
                {
                    
                    string x = (string)o["ConnectionString"];
                    string dsCatch = "Data Source=";
                    string icCatch = ";Initial Catalog=";
                    int dsCatchLen = dsCatch.Length;
                    int icCatchLen = icCatch.Length;
                    svName = x.Substring(x.IndexOf(dsCatch)+dsCatchLen,x.IndexOf(";")-x.IndexOf(dsCatch)-dsCatchLen);
                    int svNameLen = svName.Length;
                    dbName = x.Substring(x.IndexOf(icCatch)+icCatchLen);
                    
                }
                catch
                {                    
                }
            }            
        }
    }
    catch
    {
        try
        {
            dbName = (string)connjson["RemoteArtifacts"][0]["DatasetId"];
            connType = "localPowerQuery";
        }
        catch
        {
        }            
    }
    
    Connections.Add(new Connection {ServerName = svName, DatabaseName = dbName, Type = connType});

    // Custom Visuals
    try
    {
        foreach (var o in json["resourcePackages"].Children())
        {
            int resType = (int)o["resourcePackage"]["type"];
            string visualName = (string)o["resourcePackage"]["name"];
                
            if (resType == 0)
            {            
                CustomVisuals.Add( new CustomVisual {Name = visualName});
            }    
        }
    }
    catch
    {
    }

    // Report-Level Filters
    string rptFilters = json["filters"];

    try
    {
        string formattedrptfiltersJson = Newtonsoft.Json.Linq.JToken.Parse(rptFilters).ToString();
        dynamic rptFiltersJson = Newtonsoft.Json.Linq.JArray.Parse(formattedrptfiltersJson);

        foreach (var o in rptFiltersJson.Children())
        {
            string filterName = (string)o["name"];
            string filterType = (string)o["type"];
            string objectType = string.Empty;
            string objectName = string.Empty;
            string tableName = string.Empty;
            
            // Note: Add filter conditions
            try
            {
                objectName = (string)o["expression"]["Column"]["Property"];
                objectType = "Column";
                tableName = (string)o["expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
            }
            catch
            {
            }
            try
            {
                objectName = (string)o["expression"]["Measure"]["Property"];
                objectType = "Measure";
                tableName = (string)o["expression"]["Measure"]["Expression"]["SourceRef"]["Entity"];
            }
            catch
            {
            }
            try
            {
                string levelName = (string)o["expression"]["HierarchyLevel"]["Level"];
                string hierName = (string)o["expression"]["HierarchyLevel"]["Expression"]["Hierarchy"]["Hierarchy"];
                objectName = hierName + "." + levelName;
                objectType = "Hierarchy";
                tableName = (string)o["expression"]["HierarchyLevel"]["Expression"]["Hierarchy"]["Expression"]["SourceRef"]["Entity"];
            }
            catch
            {
            }
            
            ReportFilters.Add(new ReportFilter {FilterName = filterName, TableName = tableName, ObjectName = objectName, ObjectType = objectType, FilterType = filterType});
        }
    }
    catch
    {        
    }

    // Pages
    foreach (var o in json["sections"].Children())
    {
        string pageId = (string)o["name"];
        string pageName = (string)o["displayName"];
        int pageNumber = (int)o["ordinal"];
        int pageWidth = (int)o["width"];
        int pageHeight = (int)o["height"];
        int visualCount = (int)o["visualContainers"].Count;
        string pageFlt = (string)o["filters"];
        string formattedpagfltJson = Newtonsoft.Json.Linq.JToken.Parse(pageFlt).ToString();
        dynamic pageFltJson = Newtonsoft.Json.Linq.JArray.Parse(formattedpagfltJson);

        bool pageHid = false;
        string pageConfig = (string)o["config"];
        string formattedpagconfigJson = Newtonsoft.Json.Linq.JToken.Parse(pageConfig).ToString();
        dynamic pageConfigJson = Newtonsoft.Json.Linq.JObject.Parse(formattedpagconfigJson);

        try
        {
            int pageV = (int)pageConfigJson["visibility"];

            if (pageV == 1)
            {
                pageHid = true;
            }
        }
        catch
        {            
        }

        Pages.Add(new Page {Id = pageId, Name = pageName, Number = pageNumber, Width = pageWidth, Height = pageHeight, HiddenFlag = pageHid, VisualCount = visualCount });

        // Page-Level Filters
        foreach (var o2 in pageFltJson.Children())
        {
            string pgFltName = (string)o2["name"];
            string pgFltType = (string)o2["type"];
            string objType = string.Empty;
            string objName = string.Empty;
            string tblName = string.Empty;
            
            // Note: Add filter conditions
            try
            {
                objName = (string)o2["expression"]["Column"]["Property"];
                objType = "Column";
                tblName = (string)o2["expression"]["Column"]["Expression"]["SourceRef"]["Entity"];                    
            }
            catch
            {
            }
            try
            {
                objName = (string)o2["expression"]["Measure"]["Property"];
                objType = "Measure";
                tblName = (string)o2["expression"]["Measure"]["Expression"]["SourceRef"]["Entity"];                    
            }
            catch
            {
            }
            try
            {
                string levelName = (string)o2["expression"]["HierarchyLevel"]["Level"];
                string hierName = (string)o2["expression"]["HierarchyLevel"]["Expression"]["Hierarchy"]["Hierarchy"];
                objName = hierName + "." + levelName;
                objType = "Hierarchy";
                tblName = (string)o2["expression"]["HierarchyLevel"]["Expression"]["Hierarchy"]["Expression"]["SourceRef"]["Entity"];                    
            }
            catch
            {
            }

            PageFilters.Add(new PageFilter {PageId = pageId, PageName = pageName, FilterName = pgFltName, TableName = tblName, ObjectName = objName, ObjectType = objType, FilterType = pgFltType });
        }

        // Visuals
        foreach (var vc in o["visualContainers"].Children())
        {                        
            string config = (string)vc["config"];
            string formattedconfigJson = Newtonsoft.Json.Linq.JToken.Parse(config).ToString();
            dynamic configJson = Newtonsoft.Json.Linq.JObject.Parse(formattedconfigJson);
            string visualId = (string)configJson["name"];
            int cx = Convert.ToInt32(Math.Ceiling((double)vc["x"]));
            int cy = Convert.ToInt32(Math.Ceiling((double)vc["y"]));
            int cz = Convert.ToInt32(Math.Ceiling((double)vc["z"]));
            int cw = Convert.ToInt32(Math.Ceiling((double)vc["width"]));
            int ch = Convert.ToInt32(Math.Ceiling((double)vc["height"]));
            string visualType = string.Empty;
            string visualName = string.Empty;
            bool customVisualFlag = false;
            int objCount = 0;
            bool visHid = false;

            try
            {
                visualType = (string)configJson["singleVisual"]["visualType"];
            }
            catch
            {
                visualType = "visualGroup";
            }
            
            if (CustomVisuals.Exists(a => a.Name == visualType))
            {
                customVisualFlag = true;
            }

            // Visual Name
            try
            {
                visualName = (string)configJson["singleVisualGroup"]["displayName"];
            }
            catch
            {                
            }
            try
            {
                visualName = (string)configJson["singleVisual"]["vcObjects"]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"];
                visualName = visualName.Substring(1,visualName.Length-2);
            }
            catch
            {                
            }
            if (visualName.Length == 0)
            {
                visualName = visualType;
            }

            // Visual Hidden
            try
            {
                string visH = (string)configJson["singleVisual"]["display"]["mode"];

                if (visH == "hidden")
                {
                    visHid = true;
                }
            }
            catch
            {                
            }

            try
            {
                bool visH = (bool)configJson["singleVisualGroup"]["isHidden"];

                if (visH)
                {
                    visHid = true;
                }
            }
            catch
            {                
            }

            // Visual Objects
            try
            {
                objCount = configJson["singleVisual"]["prototypeQuery"]["Select"].Count;
                foreach (var o2 in configJson["singleVisual"]["prototypeQuery"]["Select"].Children())
                {
                    string objectType = string.Empty;
                    string tableName = string.Empty;
                    string objectName = string.Empty;
                    string src = string.Empty;
                    
                    try
                    {
                        objectName = (string)o2["Column"]["Property"];
                        objectType = "Column";
                        src = (string)o2["Column"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }
                    try
                    {
                        objectName = (string)o2["Measure"]["Property"];
                        objectType = "Measure";
                        src = (string)o2["Measure"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }
                    try
                    {
                        string levelName = (string)o2["HierarchyLevel"]["Level"];
                        string hierName = (string)o2["HierarchyLevel"]["Expression"]["Hierarchy"]["Hierarchy"];
                        objectName = hierName + "." + levelName;
                        objectType = "Hierarchy";
                        src = (string)o2["HierarchyLevel"]["Expression"]["Hierarchy"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }
                    try
                    {
                        objectName = (string)o2["Aggregation"]["Expression"]["Column"]["Property"];
                        objectType = "Column";
                        src = (string)o2["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }

                    foreach (var t in configJson["singleVisual"]["prototypeQuery"]["From"].Children())
                    {
                        string n = (string)t["Name"];
                        string tbl = (string)t["Entity"];

                        if (src == n)
                        {
                            tableName = tbl;
                        }
                    }
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType});
                }
            }
            catch
            {
            }
            
            Visuals.Add(new Visual {PageName = pageName, Id = visualId, Name = visualName, Type = visualType, CustomVisualFlag = customVisualFlag, HiddenFlag = visHid, X = cx, Y = cy, Z = cz, Width = cw, Height = ch, ObjectCount = objCount});
            
            // Visual Filters
            string visfilter = (string)vc["filters"];
            
            if (visfilter != null)
            {
                string formattedvisfilterJson = Newtonsoft.Json.Linq.JToken.Parse(visfilter).ToString();
                dynamic visfilterJson = Newtonsoft.Json.Linq.JArray.Parse(formattedvisfilterJson);
                                
                foreach (var o3 in visfilterJson.Children())
                {                  
                    string filterType = (string)o3["type"];
                    string objType1 = string.Empty;
                    string objName1 = string.Empty;
                    string tblName1 = string.Empty;
                    
                    // Note: Add filter conditions
                    try
                    {
                        objName1 = (string)o3["expression"]["Column"]["Property"];
                        objType1 = "Column";
                        tblName1 = (string)o3["expression"]["Column"]["Expression"]["SourceRef"]["Entity"];                    
                    }
                    catch
                    {
                    }
                    try
                    {
                        objName1 = (string)o3["expression"]["Measure"]["Property"];
                        objType1 = "Measure";
                        tblName1 = (string)o3["expression"]["Measure"]["Expression"]["SourceRef"]["Entity"];                    
                    }
                    catch
                    {
                    }
                    try
                    {
                        string levelName1 = (string)o3["expression"]["HierarchyLevel"]["Level"];
                        string hierName1 = (string)o3["expression"]["HierarchyLevel"]["Expression"]["Hierarchy"]["Hierarchy"];
                        objName1 = hierName1 + "." + levelName1;
                        objType1 = "Hierarchy";
                        tblName1 = (string)o3["expression"]["HierarchyLevel"]["Expression"]["Hierarchy"]["Expression"]["SourceRef"]["Entity"];                    
                    }
                    catch
                    {
                    }
                    VisualFilters.Add(new VisualFilter {PageName = pageName, VisualId = visualId, TableName = tblName1, ObjectName = objName1, ObjectType = objType1, FilterType = filterType });
                }
            }
        }
    }

    // Bookmarks
    try
    {
        string configEnd = (string)json["config"];
        string formattedconfigEndJson = Newtonsoft.Json.Linq.JToken.Parse(configEnd).ToString();
        dynamic configEndJson = Newtonsoft.Json.Linq.JObject.Parse(formattedconfigEndJson);

        foreach (var o in configEndJson["bookmarks"].Children())
        {
            string bName = o["displayName"];
            string bId = o["name"];
            string rptPageId = o["explorationState"]["activeSection"];
                    
            Bookmarks.Add(new Bookmark {Id = bId, Name = bName, PageId = rptPageId});
        }
    }
    catch
    {
    }

    // Add results to StringBuilders
    foreach (var x in CustomVisuals.ToList())
    {
        sb_CustomVisuals.Append(fileName + '\t' + x.Name + newline);
    }
    foreach (var x in ReportFilters.ToList())
    {
        sb_ReportFilters.Append(fileName + '\t' + x.FilterName + '\t' + x.TableName + '\t' + x.ObjectName + '\t' + x.ObjectType + '\t' + x.FilterType + newline);
    }
    foreach (var x in PageFilters.ToList())
    {
        sb_PageFilters.Append(fileName + '\t' + x.PageId + '\t' + x.PageName + '\t' + x.FilterName + '\t' + x.TableName + '\t' + x.ObjectName + '\t' + x.ObjectType + '\t' + x.FilterType + newline);
    }
    foreach (var x in VisualFilters.ToList())
    {
        sb_VisualFilters.Append(fileName + '\t' + x.PageName + '\t' + x.VisualId + '\t' + x.TableName + '\t' + x.ObjectName + '\t' + x.ObjectType + '\t' + x.FilterType + newline);
    }
    foreach (var x in VisualObjects.ToList())
    {
        sb_VisualObjects.Append(fileName + '\t' + x.PageName + '\t' + x.VisualId + '\t' + x.VisualType + '\t' + x.CustomVisualFlag + '\t' + x.TableName + '\t' + x.ObjectName + '\t' + x.ObjectType + newline);
    }
    foreach (var x in Bookmarks.ToList())
    {
        sb_Bookmarks.Append(fileName + '\t' + x.Name + '\t' + x.Id + '\t' + x.PageId + newline);
    }
    foreach (var x in Pages.ToList())
    {
        sb_Pages.Append(fileName + '\t' + x.Id + '\t' + x.Name + '\t' + x.Number + '\t' + x.Width + '\t' + x.Height + '\t' + x.HiddenFlag + '\t' + x.VisualCount + newline);
    }
    foreach (var x in Visuals.ToList())
    {
        sb_Visuals.Append(fileName + '\t' + x.PageName + '\t' + x.Id + '\t' + x.Name + '\t' + x.Type + '\t' + x.CustomVisualFlag + '\t' + x.HiddenFlag + '\t' + x.X + '\t' + x.Y + '\t' + x.Z + '\t' + x.Width + '\t' + x.Height + '\t' + x.ObjectCount + newline);
    }
    foreach (var x in Connections.ToList())
    {
        sb_Connections.Append(fileName + '\t' + x.ServerName + '\t' + x.DatabaseName + '\t' + x.Type + newline);
    }
}

// Save to text files or print out results
if (saveToFile)
{    
    System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_CustomVisuals.txt", sb_CustomVisuals.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_ReportFilters.txt", sb_ReportFilters.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_PageFilters.txt", sb_PageFilters.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_VisualFilters.txt", sb_VisualFilters.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_VisualObjects.txt", sb_VisualObjects.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_Visuals.txt", sb_Visuals.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_Bookmarks.txt", sb_Bookmarks.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_Pages.txt", sb_Pages.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_Connections.txt", sb_Connections.ToString());
}
else
{
    sb_CustomVisuals.Output();
    sb_ReportFilters.Output();
    sb_PageFilters.Output();
    sb_VisualFilters.Output();
    sb_VisualObjects.Output();
    sb_Visuals.Output();
    sb_Bookmarks.Output();
    sb_Pages.Output();
    sb_Connections.Output();  
}

}

// Classes for each object set
public class CustomVisual
{
    public string Name { get; set; }
}

public class Bookmark
{
    public string Id { get; set; }
    public string Name { get; set; }
    public string PageId { get; set; }
}

public class ReportFilter
{
    public string FilterName { get; set; }
    public string TableName { get; set; }
    public string ObjectName { get; set; }
    public string ObjectType { get; set; }
    public string FilterType { get; set; }
}

public class VisualObject
{
    public string PageName { get; set; }
    public string VisualId { get; set; }
    public string VisualType { get; set; }
    public bool CustomVisualFlag { get; set; }
    public string TableName { get; set; }
    public string ObjectName { get; set; }
    public string ObjectType { get; set; }
}

public class Visual
{
    public string PageName { get; set; }
    public string Id { get; set; }
    public string Name { get; set; }
    public string Type { get; set; }
    public bool CustomVisualFlag { get; set; }
    public bool HiddenFlag { get; set; }
    public int X { get; set; }
    public int Y { get; set; }
    public int Z { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int ObjectCount { get; set; }
}

public class VisualFilter
{
    public string PageName { get; set; }
    public string VisualId { get; set; }
    public string TableName { get; set; }
    public string ObjectName { get; set; }
    public string ObjectType { get; set; }
    public string FilterType { get; set; }
}

public class PageFilter
{
    public string PageId { get; set; }
    public string PageName { get; set; }
    public string FilterName {get; set; }
    public string TableName { get; set; }
    public string ObjectName { get; set; }
    public string ObjectType { get; set; }
    public string FilterType { get; set; }    
}

public class Page
{
    public string Id { get; set; }
    public string Name { get; set; }
    public int Number { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public bool HiddenFlag { get; set; }
    public int VisualCount {get; set; }
}

public class Connection
{
    public string ServerName { get; set; }
    public string DatabaseName { get; set; }
    public string Type { get; set; }
}

static void _() {