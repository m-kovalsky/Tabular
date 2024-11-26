#r "System.IO"
#r "System.IO.Compression.FileSystem"

using System.IO;
using System.IO.Compression;
using System.Text;

/************************************************************************************************************

    Instructions:

    1. Paste this code into the Advanced Scripting window of Tabular Editor.
    2. Create a new model in Tabular Editor (File -> New Model) (Compatibility Level does not matter).
    3. Choose a method below and follow its respective steps.

    Method 1: For a single Power BI Desktop (.pbix) or Power BI Template (.pbit) file:
        1. Enter the folder location of the file in the 'pbiFolderName' parameter.
        2. Enter the file name (including the extension) in the 'pbiFile' parameter.
    
        // You must open the model/dataset (behind the Power BI Report) in Tabular Editor in order for the following steps to work properly.
        3. Setting the 'addPersp' parameter to 'true' will create a new perspective which contains all of the objects used in the Power BI report including object dependencies.
        4. (Optional) Set the 'perspName' parameter to set the name of the perspective referenced in Step 3.

    Method 2: For looping through multiple .pbix or .pbit files within a folder:
        1. Enter the location of the folder in the 'pbiFolderName' parameter.
        2. Ensure that the pbiFile parameter is empty (should be this:  @""; ).

    Setting the 'saveToFile' parameter to 'true' will save the output text files to the respective folder.
    Setting the 'saveToFile' parameter to 'false' will generate the output within pop-up windows. 

************************************************************************************************************/

// User Parameters
string pbiFolderName = @"C:\Desktop\MyReport";
string pbiFile = @"MyReport.pbix";
bool saveToFile = true;
bool addPersp = false;
string perspName = "RptObj";

// Do not modify these parameters
string newline = Environment.NewLine;
string savePrefix = "ReportObjects";
bool singleFile = true;
bool createPersp = false;
if (pbiFile.Length == 0)
{
    singleFile = false;
}

if (singleFile && addPersp)
{
    createPersp = true;
    if (Model.Perspectives.Any(a => a.Name == perspName))
    {
        Model.Perspectives[perspName].Delete();
    }
    Model.AddPerspective(perspName);
}

List<string> FileList = new List<string>();

var sb_CustomVisuals = new System.Text.StringBuilder();
sb_CustomVisuals.Append("ReportName" + '\t' + "Name" + '\t' + "ReportDate" + newline);

var sb_ReportFilters = new System.Text.StringBuilder();
sb_ReportFilters.Append("ReportName" + '\t' + "DisplayName" + '\t' + "TableName" + '\t' + "ObjectName" + '\t' + "ObjectType" + '\t' + "FilterType"  + '\t' + "HiddenFilter"  + '\t' + "LockedFilter"  + '\t' + "AppliedFilterVersion"  + '\t' + "ReportDate" + newline);

var sb_VisualObjects = new System.Text.StringBuilder();
sb_VisualObjects.Append("ReportName" + '\t' + "PageName" + '\t' + "VisualId" + '\t' + "VisualType" + '\t' + "AppliedFilterVersion" + '\t' + "CustomVisualFlag" + '\t' + "TableName" + '\t' + "ObjectName" + '\t' + "ObjectType" + '\t' + "Source" + '\t' + "DisplayName" + '\t' + "ReportDate" + newline);

var sb_VisualFilters = new System.Text.StringBuilder();
sb_VisualFilters.Append("ReportName" + '\t' + "PageName" + '\t' + "VisualId" + '\t' + "TableName" + '\t' + "ObjectName" + '\t' + "ObjectType" + '\t' + "FilterType" + '\t' + "HiddenFilter"  + '\t' + "LockedFilter"  + '\t' + "AppliedFilterVersion"  + '\t' + "DisplayName" + '\t' + "ReportDate" + newline);

var sb_PageFilters = new System.Text.StringBuilder();
sb_PageFilters.Append("ReportName" + '\t' + "PageId" + '\t' + "PageName" + '\t' + "DisplayName" + '\t' + "TableName" + '\t' + "ObjectName" + '\t' + "ObjectType" + '\t' + "FilterType" + '\t' + "HiddenFilter"  + '\t' + "LockedFilter"  + '\t' + "AppliedFilterVersion"  + '\t' + "ReportDate" + newline);

var sb_Bookmarks = new System.Text.StringBuilder();
sb_Bookmarks.Append("ReportName" + '\t' + "Name" + '\t' + "Id" + '\t' + "PageId" + '\t' + "VisualId" + '\t' + "VisualHiddenFlag" + '\t' + "ReportDate" + newline);

var sb_Pages = new System.Text.StringBuilder();
sb_Pages.Append("ReportName" + '\t' + "Id" + '\t' + "Name" + '\t' + "Number" + '\t' + "Width" + '\t' + "Height" + '\t' + "HiddenFlag" + '\t' + "VisualCount" + '\t' + "BackgroundImage" + '\t' + "WallpaperImage" + '\t' + "Type" + '\t' + "ReportDate" + newline);

var sb_Visuals = new System.Text.StringBuilder();
sb_Visuals.Append("ReportName" + '\t' + "PageName" + '\t' + "Id" + '\t' + "Name" + '\t' + "Type" + '\t' + "CustomVisualFlag" + '\t' + "HiddenFlag" + '\t' + "X" + '\t' + "Y" + '\t' + "Z" + '\t' + "Width" + '\t' + "Height" + '\t' + "ObjectCount" + '\t' + "ShowItemsNoDataFlag" + '\t' + "SlicerType" + '\t' + "ParentGroup" + '\t' + "ReportDate" + newline);

var sb_Connections = new System.Text.StringBuilder();
sb_Connections.Append("ReportName" + '\t' + "ServerName" + '\t' + "DatabaseName" + '\t' + "ReportID" + '\t' + "Type" + '\t' + "ReportDate" + newline);

var sb_VisualInteractions = new System.Text.StringBuilder();
sb_VisualInteractions.Append("ReportName" + '\t' + "PageName" + '\t' + "SourceVisualID" + '\t' + "TargetVisualID" + '\t' + "TypeID" + '\t' + "Type" + '\t' + "ReportDate" + newline);

if (pbiFile.Length > 0 && pbiFolderName.Length == 0)
{
    Error("If specifying the 'pbiFile' you must also specify the 'pbiFolderName'.");
    return;
}    
else if (pbiFile.Length == 0 && pbiFolderName.Length > 0)
{
    foreach (var x in System.IO.Directory.GetFiles(pbiFolderName, "*.pbi*"))
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
    var VisualInteractions = new List<VisualInteraction>();
    string fileExt = Path.GetExtension(rpt);
    string ReportName = Path.GetFileNameWithoutExtension(rpt);
    string ReportDate = DateTime.Now.ToString("yyyy-MM-dd");
    string folderName = Path.GetDirectoryName(rpt) + @"\";
    string zipPath = folderName + ReportName + ".zip";
    string unzipPath = folderName + ReportName;

if (!(fileExt == ".pbix" || fileExt == ".pbit"))
{
    Error("'" + rpt + "' is not a valid file. File(s) must be a valid .pbix or .pbit.");
    return;
}

bool extractionSucceeded = false;

try
{
    // Make a copy of a pbi and turn it into a zip file
    File.Copy(rpt, zipPath);
    // Unzip file
    System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, unzipPath);
    extractionSucceeded = true; // Set flag to true if extraction is successful
}
catch (System.IO.PathTooLongException)
{
    // Log or handle the PathTooLongException if needed
}
catch (System.IO.DirectoryNotFoundException)
{
    // Log or handle the DirectoryNotFoundException if needed
}
catch (System.IO.InvalidDataException ex)
{

}
catch (Exception ex)
{
    
}
finally
    {
        // Delete zip file
        File.Delete(zipPath);
    }




    // Layout file
   
    string layoutPath = unzipPath + @"\Report\Layout";
    string jsonFilePath = Path.ChangeExtension(layoutPath, ".json");
    
        if (!File.Exists(layoutPath))
    {
        continue; // If Layout file is not found, continue to the next file
    }
    
    File.Move(layoutPath, jsonFilePath); 

    string unformattedJson = File.ReadAllText(jsonFilePath,System.Text.UnicodeEncoding.Unicode);
    string formattedJson = Newtonsoft.Json.Linq.JToken.Parse(unformattedJson).ToString();
    dynamic json = Newtonsoft.Json.Linq.JObject.Parse(formattedJson);
    
    







// ****Connections file****








    string svName = string.Empty;
    string dbName = string.Empty;
    string datetoday = string.Empty;
    string rptId = string.Empty;
    string connType = string.Empty;
    string connPath = unzipPath + @"\Connections";
    if (File.Exists(connPath))
    {        
        string jsonconnFilePath = Path.ChangeExtension(connPath, ".json");
        File.Move(connPath, jsonconnFilePath); 

        string unformattedconnJson = File.ReadAllText(jsonconnFilePath,System.Text.Encoding.UTF8);
        string formattedconnJson = Newtonsoft.Json.Linq.JToken.Parse(unformattedconnJson).ToString();
        dynamic connjson = Newtonsoft.Json.Linq.JObject.Parse(formattedconnJson);
        
        // Connection info
try
{
    foreach (var o in connjson["Connections"].Children())
    {
        connType = (string)o["ConnectionType"];

        // For specific connection types
        if (connType == "pbiServiceLive" || connType == "pbiServiceXmlaStyleLive")
        {
            try
            {
                dbName = (string)connjson["RemoteArtifacts"][0]["DatasetId"];
                rptId = (string)connjson["RemoteArtifacts"][0]["ReportId"];
            }
            catch
            {
            }
        }
        else
        {
            // For other connection types
            try
            {
                dbName = (string)o["PbiModelDatabaseName"];
            }
            catch
            {
            }

            try
            {
                string x = (string)o["ConnectionString"];
                string dsCatch = "Data Source=";
                string icCatch = ";Initial Catalog=";
                int dsCatchLen = dsCatch.Length;
                int icCatchLen = icCatch.Length;

                svName = x.Substring(x.IndexOf(dsCatch) + dsCatchLen, x.IndexOf(";") - x.IndexOf(dsCatch) - dsCatchLen);
                dbName = x.Substring(x.IndexOf(icCatch) + icCatchLen);
            }
            catch
            {
            }
        }
    }
}
catch
{
    // Fallback for localPowerQuery
    try
    {
        dbName = (string)connjson["RemoteArtifacts"][0]["DatasetId"];
        rptId = (string)connjson["RemoteArtifacts"][0]["ReportId"];
        connType = "localPowerQuery";
    }
    catch
    {
    }
}
        Connections.Add(new Connection {ServerName = svName, DatabaseName = dbName, Type = connType, ReportID = rptId});        
    }
    
    //Delete previously created folder
    try
    {
        Directory.Delete(folderName + ReportName,true);
    }
    catch
    {
    }

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
            else if (resType == 1)
            {
                try
                {
                    // Create list of images used
                    foreach (var o2 in o["resourcePackage"]["items"].Children())
                    {
                        string imageName = o2["path"];
                    }
                }
                catch
                {
                }
            }
        }
    }
    catch
    {
    }








    // ****Report-Level Filters****








    string rptFilters = json["filters"];

    try
    {
        string formattedrptfiltersJson = Newtonsoft.Json.Linq.JToken.Parse(rptFilters).ToString();
        dynamic rptFiltersJson = Newtonsoft.Json.Linq.JArray.Parse(formattedrptfiltersJson);

        foreach (var o in rptFiltersJson.Children())
        {
            string displayName = (string)o["displayName"];
            string filterType = (string)o["type"];
            string hiddenfilter = (string)o["isHiddenInViewMode"];
            string lockedfilter = (string)o["isLockedInViewMode"];
            string objectType = string.Empty;
            string objectName = string.Empty;
            string tableName = string.Empty;
            string appliedfilterversion = string.Empty;
            
            // Note: Add filter conditions

try
{
    if (o != null && o.ContainsKey("filter") && o["filter"].ContainsKey("Version"))
    {
        appliedfilterversion = o["filter"]["Version"].ToString();
    }
}
catch (Exception ex)
{
}

            try
            {
                objectName = (string)o["expression"]["Column"]["Property"];
                objectType = "Column";
                tableName = (string)o["expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                
                if (createPersp)
                {
                    try
                    {
                        Model.Tables[tableName].Columns[objectName].InPerspective[perspName] = true;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
            try
            {
                objectName = (string)o["expression"]["Measure"]["Property"];
                objectType = "Measure";
                tableName = (string)o["expression"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                
                if (createPersp)
                {
                    try
                    {
                        Model.Tables[tableName].Measures[objectName].InPerspective[perspName] = true;
                    }
                    catch
                    {
                    }
                }
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
                
                if (createPersp)
                {
                    try
                    {
                        Model.Tables[tableName].Hierarchies[hierName].InPerspective[perspName] = true;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }           
            ReportFilters.Add(new ReportFilter {displayName = displayName, TableName = tableName, ObjectName = objectName, ObjectType = objectType, FilterType = filterType, HiddenFilter = hiddenfilter, LockedFilter = lockedfilter, AppliedFilterVersion = appliedfilterversion});
        }
    }
    catch
    {        
    }









// ****Pages****








foreach (var o in json["sections"].Children())
{
    string pageId = (string)o["name"];
    
    // Check if "displayName" exists in the JSON object before assigning
    string pageName = o["displayName"] != null ? (string)o["displayName"] : ReportName;

    string pageFlt = (string)o["filters"];
    int pageNumber = 0;
    try
    {
        pageNumber = (int)o["ordinal"];
    }
    catch
    {
    }

    int pageWidth = (int)o["width"];
    int pageHeight = (int)o["height"];
    int visualCount = (int)o["visualContainers"].Count;
    string pageBkgrd = "";
    string pageWall = "";

    // Use ReportName if pageName is null, empty, or whitespace
    if (string.IsNullOrWhiteSpace(pageName))
    {
        pageName = ReportName;
    }

string formattedpagfltJson = "";
if (!string.IsNullOrEmpty(pageFlt))
{
    formattedpagfltJson = Newtonsoft.Json.Linq.JToken.Parse(pageFlt).ToString();
}
dynamic pageFltJson = string.IsNullOrEmpty(formattedpagfltJson) 
    ? new Newtonsoft.Json.Linq.JArray() 
    : Newtonsoft.Json.Linq.JArray.Parse(formattedpagfltJson);

int displayOpt = (int)o["displayOption"];
bool pageHid = false;
string pageType = string.Empty;
string pageConfig = (string)o["config"];

string formattedpagconfigJson = "";
if (!string.IsNullOrEmpty(pageConfig))
{
    formattedpagconfigJson = Newtonsoft.Json.Linq.JToken.Parse(pageConfig).ToString();
}
dynamic pageConfigJson = string.IsNullOrEmpty(formattedpagconfigJson) 
    ? new Newtonsoft.Json.Linq.JObject() 
    : Newtonsoft.Json.Linq.JObject.Parse(formattedpagconfigJson);

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
        








        // ****Visual Interactions****








        try
        {
            foreach (var o2 in pageConfigJson["relationships"].Children())
            {
                string sourceViz = (string)o2["source"];
                string targetViz = (string)o2["target"];
                int typeID = (int)o2["type"];
                string[] typeAr = {"blank", "Filter", "Highlight", "None"};
                string type = typeAr[typeID];
                
                VisualInteractions.Add(new VisualInteraction {PageName = pageName, SourceVisualID = sourceViz, TargetVisualID = targetViz, TypeID = typeID, Type = type});
            }
        }
        catch
        {
        }
        
        // Page Background
        try
        {
            pageBkgrd = pageConfigJson["objects"]["background"][0]["properties"]["image"]["image"]["url"]["expr"]["ResourcePackageItem"]["ItemName"];
        }
        catch
        {
        }
        
        // Page Wallpaper
        try
        {
            pageWall = pageConfigJson["objects"]["outspace"][0]["properties"]["image"]["image"]["url"]["expr"]["ResourcePackageItem"]["ItemName"];
        }
        catch
        {
        }
        
        // Page Type
        if (displayOpt == 3 && pageWidth == 320 && pageHeight == 240)
        {
            pageType = "Tooltip";
        }
        else if (pageWidth == 816 && pageHeight == 1056)
        {
            pageType = "Letter";
        }
        else if (pageWidth == 960 && pageHeight == 720)            
        {
            pageType = "4:3";
        }
        else if (pageWidth == 1280 && pageHeight == 720)
        {
            pageType = "16:9";
        }
        else
        {
            pageType = "Custom";
        }
        
        Pages.Add(new Page {Id = pageId, Name = pageName, Number = pageNumber, Width = pageWidth, Height = pageHeight, HiddenFlag = pageHid, VisualCount = visualCount, BackgroundImage = pageBkgrd, WallpaperImage = pageWall, Type = pageType });

        







// ****Page-Level Filters****








        foreach (var o2 in pageFltJson.Children())
        {
            string displayName = (string)o2["displayName"];
            string pgFltType = (string)o2["type"];
            string pghiddenfilter = (string)o2["isHiddenInViewMode"];
            string pglockedfilter = (string)o2["isLockedInViewMode"];
            string objType = string.Empty;
            string objName = string.Empty;
            string tblName = string.Empty;
            string pgappliedfilterversion = string.Empty;
            
            // Note: Add filter conditions

try
{
    if (o2 != null && o2.ContainsKey("filter") && o2["filter"].ContainsKey("Version"))
    {
        pgappliedfilterversion = o2["filter"]["Version"].ToString();
    }
}
catch (Exception ex)
{
}
            try
            {
                objName = (string)o2["expression"]["Column"]["Property"];
                objType = "Column";
                tblName = (string)o2["expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                
                if (createPersp)
                {
                    try
                    {
                        Model.Tables[tblName].Columns[objName].InPerspective[perspName] = true;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
            try
            {
                objName = (string)o2["expression"]["Measure"]["Property"];
                objType = "Measure";
                tblName = (string)o2["expression"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                
                if (createPersp)
                {
                    try
                    {
                        Model.Tables[tblName].Measures[objName].InPerspective[perspName] = true;
                    }
                    catch
                    {
                    }
                }
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
                
                if (createPersp)
                {
                    try
                    {
                        Model.Tables[tblName].Hierarchies[hierName].InPerspective[perspName] = true;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            PageFilters.Add(new PageFilter {PageId = pageId, PageName = pageName, displayName = displayName, TableName = tblName, ObjectName = objName, ObjectType = objType, FilterType = pgFltType, HiddenFilter = pghiddenfilter, LockedFilter = pglockedfilter, AppliedFilterVersion = pgappliedfilterversion });
        }








        // ****Visuals****








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
            string displayName2 = (string)vc["displayName"];
            string visualType = string.Empty;
            string visualName = string.Empty;
            bool customVisualFlag = false;
            int objCount = 0;
            bool visHid = false;
            bool showItemsNoData = false;
            string slicerType = "N/A";
            string parentGroup = string.Empty;
            
            try
            {
                parentGroup = (string)configJson["parentGroupName"];
            }
            catch
            {
            }         
            
        // Show Items With No Data
    try
    {
            string sInd = (string)configJson["singleVisual"]["showAllRoles"][0];
    
        // Check for "Values", "Rows", or "Columns"
        if (sInd == "Values" || sInd == "Rows" || sInd == "Columns")
        {
            showItemsNoData = true;
        }
    }
    catch
    {
    }

    // Determine Visual Type
    try
    {
        visualType = (string)configJson["singleVisual"]["visualType"];
    }
    catch
    {
        visualType = "visualGroup";
    }

    // Check if the visual type is a custom visual
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

            if (visualType == "slicer")
            {
                try
                {
                    string sT = (string)configJson["singleVisual"]["objects"]["data"][0]["properties"]["mode"]["expr"]["Literal"]["Value"];

                    if (sT == "'Basic'")
                    {
                        slicerType = "List";
                    }
                    else if (sT == "'Dropdown'")
                    {
                        slicerType = "Dropdown";
                    }
                }
                catch
                {                    
                }
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









            // ****Visual Objects****








            try
            {
                objCount = configJson["singleVisual"]["prototypeQuery"]["Select"].Count;
                foreach (var o2 in configJson["singleVisual"]["prototypeQuery"]["Select"].Children())
                {
                    string objectType = string.Empty;
                    string tableName = string.Empty;
                    string objectName = string.Empty;
                    string src = string.Empty;
                    bool isSpark = false;
                    string sourceLabel = "Standard";
                    string displayName = string.Empty;
            string appliedFilterVersion = string.Empty;
        try
        {
            string objectRef = (string)o2["Name"];
            displayName = (string)configJson["singleVisual"]["columnProperties"][objectRef]["displayName"];
        }
        catch
        {
            // Handle the case where displayName is not found
        }
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
                        objectName = (string)o2["Arithmetic"]["Left"]["Aggregation"]["Expression"]["Column"]["Property"];
                        objectType = "Column";
                        src = (string)o2["Arithmetic"]["Left"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }   
                    try
                    {
                        objectName = (string)o2["Arithmetic"]["Right"]["Aggregation"]["Expression"]["Column"]["Property"];
                        objectType = "Column";
                        src = (string)o2["Arithmetic"]["Right"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }   
                    try
                    {
                        objectName = (string)o2["Arithmetic"]["Left"]["Aggregation"]["Expression"]["Measure"]["Property"];
                        objectType = "Measure";
                        src = (string)o2["Arithmetic"]["Left"]["Aggregation"]["Expression"]["Measure"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }   
                    try
                    {
                        objectName = (string)o2["Arithmetic"]["Right"]["Aggregation"]["Expression"]["Measure"]["Property"];
                        objectType = "Measure";
                        src = (string)o2["Arithmetic"]["Right"]["Aggregation"]["Expression"]["Measure"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }                      
                     try
                    {
                        objectName = (string)o2["Arithmetic"]["Right"]["Measure"]["Property"];
                        objectType = "Measure";
                        src = (string)o2["Arithmetic"]["Right"]["Measure"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }
                    try
                    {
                        objectName = (string)o2["Arithmetic"]["Left"]["Column"]["Property"];
                        objectType = "Column";
                        src = (string)o2["Arithmetic"]["Left"]["Column"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }
                     try
                    {
                        objectName = (string)o2["Arithmetic"]["Left"]["Measure"]["Property"];
                        objectType = "Measure";
                        src = (string)o2["Arithmetic"]["Left"]["Measure"]["Expression"]["SourceRef"]["Source"];
                    }
                    catch
                    {
                    }
                    try
                    {
                        objectName = (string)o2["Arithmetic"]["Right"]["Column"]["Property"];
                        objectType = "Column";
                        src = (string)o2["Arithmetic"]["Right"]["Column"]["Expression"]["SourceRef"]["Source"];
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
                    try
                    {
                        objectName = (string)o2["Aggregation"]["Expression"]["Measure"]["Property"];
                        objectType = "Measure";
                        src = (string)o2["Aggregation"]["Expression"]["Measure"]["Expression"]["SourceRef"]["Source"];
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
                    // Sparklines
                    try
                    {
                        objectName = (string)o2["SparklineData"]["Measure"]["Measure"]["Property"];
                        objectType = "Measure";
                        src = (string)o2["SparklineData"]["Measure"]["Measure"]["Expression"]["SourceRef"]["Source"];
                        isSpark = true;
                    }
                    catch
                    {
                    }
                    try
                    {
                        objectName = (string)o2["SparklineData"]["Measure"]["Aggregation"]["Expression"]["Column"]["Property"];
                        objectType = "Column";
                        src = (string)o2["SparklineData"]["Measure"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Source"];
                        isSpark = true;
                    }
                    catch
                    {
                    }     
                    
                    try
                    {

                appliedFilterVersion = (string)configJson["singleVisual"]["objects"]["general"][0]["properties"]["filter"]["filter"]["Version"];
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
                    
                    if (createPersp)
                    {
                        if (objectType == "Column")
                        {
                            try
                            {
                                Model.Tables[tableName].Columns[objectName].InPerspective[perspName] = true;
                            }
                            catch
                            {
                            }
                        }
                        else if (objectType == "Measure")
                        {
                            try
                            {
                                Model.Tables[tableName].Measures[objectName].InPerspective[perspName] = true;
                            }
                            catch
                            {
                            }
                        }
                        else if (objectType == "Hierarchy")
                        {
                            try
                            {
                                Model.Tables[tableName].Hierarchies[objectName].InPerspective[perspName] = true;
                            }
                            catch
                            {
                            }
                        }
                    }
                    
                    if (isSpark)
                    {
                        sourceLabel = "Sparkline";
                    }
                        
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, VisualType = visualType, AppliedFilterVersion = appliedFilterVersion, displayName = displayName, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sourceLabel});
                    
                    if (isSpark)
                    {
                        try
                        {
                            objectName = (string)o2["SparklineData"]["Groupings"][0]["Column"]["Property"];
                            objectType = "Column";
                            src = (string)o2["SparklineData"]["Groupings"][0]["Column"]["Expression"]["SourceRef"]["Source"];
                            isSpark = true;
                            
                            foreach (var t in configJson["singleVisual"]["prototypeQuery"]["From"].Children())
                            {
                                string n = (string)t["Name"];
                                string tbl = (string)t["Entity"];

                                if (src == n)
                                {
                                    tableName = tbl;
                                }
                            }
                                                        
                            VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, VisualType = visualType, displayName = displayName, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sourceLabel});
                            
                            if (createPersp)
                            {
                                if (objectType == "Column")
                                {
                                    try
                                    {
                                        Model.Tables[tableName].Columns[objectName].InPerspective[perspName] = true;
                                    }
                                    catch
                                    {
                                    }
                                }
                                else if (objectType == "Measure")
                                {
                                    try
                                    {
                                        Model.Tables[tableName].Measures[objectName].InPerspective[perspName] = true;
                                    }
                                    catch
                                    {
                                    }
                                }
                                else if (objectType == "Hierarchy")
                                {
                                    try
                                    {
                                        Model.Tables[tableName].Hierarchies[objectName].InPerspective[perspName] = true;
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                        }
                        catch
                        {
                        }
                    }
                }
            }
            catch
            {
            }
            
            // VisualObjects in Labels
            try
            {
                string sc = "Label";

                // Gradient
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["labels"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Rules
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["labels"].Children())
                    {
                        string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field value (Column)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["labels"].Children())
                    {
                        string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field Value (Measure)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["labels"].Children())
                    {
                        string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }

            // VisualObjects in CategoryAxis
            try
            {
                string sc = "X Axis Color";

                // Gradient
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Rules
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field value (Column)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field Value (Measure)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string tableName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }

            // VisualObjects in CategoryAxis Title Color
            try
            {
                string sc = "X Axis Title Color";

                // Gradient
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Rules
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field value (Column)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field Value (Measure)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string tableName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }

            // VisualObjects in ValueAxis Min
            try
            {
                string sc = "Y Axis Minimum";

                // Field value (Column)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["start"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["start"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field Value (Measure)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string tableName = (string)o2["properties"]["start"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["start"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }

            // VisualObjects in ValueAxis Max
            try
            {
                string sc = "Y Axis Minimum";

                // Field value (Column)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["end"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["end"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field Value (Measure)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryAxis"].Children())
                    {
                        string tableName = (string)o2["properties"]["end"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["end"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }

            try
            {
                string sc = "Y Axis Color";

                // Gradient
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["valueAxis"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Rules
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["valueAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field value (Column)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["valueAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field Value (Measure)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["valueAxis"].Children())
                    {
                        string tableName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["labelColor"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }

            // VisualObjects in ValueAxis Title Color
            try
            {
                string sc = "Y Axis Title Color";

                // Gradient
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["valueAxis"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Rules
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["valueAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field value (Column)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["valueAxis"].Children())
                    {
                        string objectName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }

                // Field Value (Measure)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["valueAxis"].Children())
                    {
                        string tableName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["titleColor"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }            

            // VisualObjects in Category Labels
            try
            {
                string sc = "Category Label";
                // Gradient
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryLabels"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
                // Rules
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryLabels"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
                // Field Value - Column
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryLabels"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expresssion"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
                // Field Value - Measure
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["objects"]["categoryLabels"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }
            // Objects in Text (???)
            try
            {
                string sc = "Text";
                foreach (var o2 in configJson["singleVisual"]["objects"]["text"].Children())
                {
                    try
                    {
                        string tableName = (string)o2["properties"]["text"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];                    
                        string objectName = (string)o2["properties"]["text"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }                    
                }
            }
            catch
            {
            }
       
            // Conditional Formatting (Background Color)
            try
            {
                string sc = "Conditional Formatting (Background Color)";
                foreach (var o2 in configJson["singleVisual"]["objects"]["values"].Children())
                {
                    // Gradient
                    try
                    {
                        string objectName = (string)o2["properties"]["backColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["backColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }

                    // Rules
                    try
                    {
                        string objectName = (string)o2["properties"]["backColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["backColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";                    
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }

                    // Field value (Column)
                    try
                    {
                        string objectName = (string)o2["properties"]["backColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["backColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";                    
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }
                    //Field value (Measure)
                    try
                    {
                        string tableName = (string)o2["properties"]["backColor"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["backColor"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                        
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
            // Conditional Formatting (Font Color)
            try
            {
                string sc = "Conditional Formatting (Font Color)";
                foreach (var o2 in configJson["singleVisual"]["objects"]["values"].Children())
                {
                    // Gradient
                    try
                    {
                        string objectName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";                    
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }

                    // Rules
                    try
                    {
                        string objectName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }

                    // Field value
                    try
                    {
                        string objectName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            // Conditional Formatting (Icon)
            try
            {
                string sc = "Conditional Formatting (Icon)";
                foreach (var o2 in configJson["singleVisual"]["objects"]["values"].Children())
                {
                    // Rules
                    try
                    {
                        string objectName = (string)o2["properties"]["icon"]["value"]["expr"]["Conditional"]["Cases"][0]["Condition"]["Comparison"]["Left"]["Measure"]["Property"];                    
                        string tableName = (string)o2["properties"]["icon"]["value"]["expr"]["Conditional"]["Cases"][0]["Condition"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }

                    // Field value
                    try
                    {
                        string objectName = (string)o2["properties"]["icon"]["value"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["icon"]["value"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            // Conditional Formatting (WebURL)
            try
            {
                string sc = "Conditional Formatting (WebURL)";
                foreach (var o2 in configJson["singleVisual"]["objects"]["values"].Children())
                {
                    // Field value
                    try
                    {
                        string objectName = (string)o2["properties"]["webURL"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string tableName = (string)o2["properties"]["webURL"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
            
            // VisualObjects in Title Text
            try
            {                
                string sc = "Title Text";
                // Field Value (Column)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {                    
                        // labels
                        string tableName = (string)o2["properties"]["text"]["expr"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["text"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }
                // Field Value (Measure)
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["text"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["text"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {                    
                }
            }
            catch
            {
            }
            // VisualObjects in Title Font Color
            try
            {
                string sc = "Title Font Color";
                // Gradient
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
                // Rules
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
                // Field Value - Column
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expresssion"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
                // Field Value - Measure
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["fontColor"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, displayName = displayName, VisualId = visualId, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }
            // VisualObjects in Title Background
            try
            {
                string sc = "Title Background";
                // Gradient
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["background"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["background"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
                // Rules
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["background"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["background"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
                // Field Value - Column
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["background"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expresssion"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["background"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Column";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
                // Field Value - Measure
                try
                {
                    foreach (var o2 in configJson["singleVisual"]["vcObjects"]["title"].Children())
                    {
                        // labels
                        string tableName = (string)o2["properties"]["background"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        string objectName = (string)o2["properties"]["background"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                        string objectType = "Measure";
                        
                        VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }
            // VisualObjects in Background (Gradient)
            try
            {
                string sc = "Background";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["background"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Measure";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }
            // VisualObjects in Background (Rules)
            try
            {
                string sc = "Background";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["background"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Measure";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }
            // VisualObjects in Background (Field Value - Column)
            try
            {
                string sc = "Background";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["background"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expresssion"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Column";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName,  VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }
            // VisualObjects in Background (Field Value - Measure)
            try
            {
                string sc = "Background";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["background"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Measure";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }

            // VisualObjects in Border (Gradient)
            try
            {
                string sc = "Border";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["border"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Measure";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }
            // VisualObjects in Border (Rules)
            try
            {
                string sc = "Border";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["border"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Measure";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }
            // VisualObjects in Border (Field Value - Column)
            try
            {
                string sc = "Border";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["border"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expresssion"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                    string displayName = (string)o2["displayName"];
                    string objectType = "Column";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }
            // VisualObjects in Border (Field Value - Measure)
            try
            {
                string sc = "Border";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["border"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Measure";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }

            // VisualObjects in Drop Shadow (Gradient)
            try
            {
                string sc = "Drop Shadow";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["dropShadow"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["FillRule"]["Input"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Measure";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }
            // VisualObjects in Drop Shadow (Rules)
            try
            {
                string sc = "Drop Shadow";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["dropShadow"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Conditional"]["Cases"][0]["Condition"]["And"]["Left"]["Comparison"]["Left"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Measure";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }
            // VisualObjects in Drop Shadow (Field Value - Column)
            try
            {
                string sc = "Drop Shadow";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["dropShadow"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Expresssion"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Aggregation"]["Expression"]["Column"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Column";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }
            // VisualObjects in Drop Shadow (Field Value - Measure)
            try
            {
                string sc = "Drop Shadow";
                foreach (var o2 in configJson["singleVisual"]["vcObjects"]["dropShadow"].Children())
                {
                    // labels
                    string tableName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                    string objectName = (string)o2["properties"]["color"]["solid"]["color"]["expr"]["Measure"]["Property"];
                        string displayName = (string)o2["displayName"];
                    string objectType = "Measure";
                    
                    VisualObjects.Add(new VisualObject {PageName = pageName, VisualId = visualId, displayName = displayName, VisualType = visualType, CustomVisualFlag = customVisualFlag, ObjectName = objectName, TableName = tableName, ObjectType = objectType, Source = sc});
                }
            }
            catch
            {
            }

            Visuals.Add(new Visual {PageName = pageName, Id = visualId, Name = visualName, Type = visualType, CustomVisualFlag = customVisualFlag, HiddenFlag = visHid, X = cx, Y = cy, Z = cz, Width = cw, Height = ch, ObjectCount = objCount, ShowItemsNoDataFlag = showItemsNoData, SlicerType = slicerType, ParentGroup = parentGroup });
            
            







// ****Visual Filters****








            string visfilter = (string)vc["filters"];
            
            if (visfilter != null)
            {
                string formattedvisfilterJson = Newtonsoft.Json.Linq.JToken.Parse(visfilter).ToString();
                dynamic visfilterJson = Newtonsoft.Json.Linq.JArray.Parse(formattedvisfilterJson);
                                
                foreach (var o3 in visfilterJson.Children())
                {                  
                    string displayName = (string)o3["displayName"];
                    string filterType = (string)o3["type"];
                    string hiddenfilter = (string)o3["isHiddenInViewMode"];
                    string lockedfilter = (string)o3["isLockedInViewMode"];
                    string objType1 = string.Empty;
                    string objName1 = string.Empty;
                    string tblName1 = string.Empty;
            string appliedfilterversion = string.Empty;
            
            // Note: Add filter conditions

try
{
    if (o3 != null && o3.ContainsKey("filter") && o3["filter"].ContainsKey("Version"))
    {
        appliedfilterversion = o3["filter"]["Version"].ToString();
    }
}
catch (Exception ex)
{
}
                    try
                    {
                        objName1 = (string)o3["expression"]["Column"]["Property"];
                        objType1 = "Column";
                        tblName1 = (string)o3["expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
                        
                        if (createPersp)
                        {
                            try
                            {
                                Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                            }
                            catch
                            {
                            }
                        }
                    }
                    catch
                    {
                    }
                    try
                    {
                        objName1 = (string)o3["expression"]["Measure"]["Property"];
                        objType1 = "Measure";
                        tblName1 = (string)o3["expression"]["Measure"]["Expression"]["SourceRef"]["Entity"];
                        
                        if (createPersp)
                        {
                            try
                            {
                                Model.Tables[tblName1].Measures[objName1].InPerspective[perspName] = true;
                            }
                            catch
                            {
                            }
                        }
                    }
                    catch
                    {
                    }
                    
                            try
        {
            objName1 = (string)o3["expression"]["Aggregation"]["Expression"]["Column"]["Property"];
            tblName1 = (string)o3["expression"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Column";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
try
        {
            objName1 = (string)o3["expression"]["Aggregation"]["Expression"]["Measure"]["Property"];
            tblName1 = (string)o3["expression"]["Aggregation"]["Expression"]["Measure"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Measure";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
                             try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Left"]["Measure"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Left"]["Measure"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Measure";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
                             try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Left"]["Column"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Left"]["Column"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Column";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Right"]["Measure"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Right"]["Measure"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Measure";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Right"]["Column"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Right"]["Column"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Column";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Left"]["Aggregation"]["Expression"]["Column"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Left"]["Aggregation"]["Expression"]["Column"]["Expresssion"]["SourceRef"]["Entity"];
            objType1 = "Column";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Right"]["Aggregation"]["Expression"]["Column"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Right"]["Aggregation"]["Expression"]["Column"]["Expresssion"]["SourceRef"]["Entity"];
            objType1 = "Column";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Right"]["ScopedEval"]["Expression"]["Aggregation"]["Expression"]["Column"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Right"]["ScopedEval"]["Expression"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Column";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Left"]["ScopedEval"]["Expression"]["Aggregation"]["Expression"]["Column"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Left"]["ScopedEval"]["Expression"]["Aggregation"]["Expression"]["Column"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Column";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Right"]["ScopedEval"]["Expression"]["Aggregation"]["Expression"]["Measure"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Right"]["ScopedEval"]["Expression"]["Aggregation"]["Expression"]["Measure"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Measure";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Left"]["ScopedEval"]["Expression"]["Aggregation"]["Expression"]["Measure"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Left"]["ScopedEval"]["Expression"]["Aggregation"]["Expression"]["Measure"]["Expression"]["SourceRef"]["Entity"];
            objType1 = "Measure";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Left"]["Aggregation"]["Expression"]["Measure"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Left"]["Aggregation"]["Expression"]["Measure"]["Expresssion"]["SourceRef"]["Entity"];
            objType1 = "Measure";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
        }
        catch
        {
        }
         try
        {
            objName1 = (string)o3["expression"]["Arithmetic"]["Right"]["Aggregation"]["Expression"]["Measure"]["Property"];
            tblName1 = (string)o3["expression"]["Arithmetic"]["Right"]["Aggregation"]["Expression"]["Measure"]["Expresssion"]["SourceRef"]["Entity"];
            objType1 = "Measure";
            
            if (createPersp)
            {
                try
                {
                    Model.Tables[tblName1].Columns[objName1].InPerspective[perspName] = true;
                }
                catch
                {
                    // Handle exception
                }
            }
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
                        
                        if (createPersp)
                        {
                            try
                            {
                                Model.Tables[tblName1].Hierarchies[hierName1].InPerspective[perspName] = true;
                            }
                            catch
                            {
                            }
                        }
                    }
                    catch
                    {
                    }
                    VisualFilters.Add(new VisualFilter {PageName = pageName, VisualId = visualId, displayName = displayName, TableName = tblName1, ObjectName = objName1, ObjectType = objType1, FilterType = filterType, HiddenFilter = hiddenfilter, LockedFilter = lockedfilter, AppliedFilterVersion = appliedfilterversion });
                }
            }
        }
    }

    







// ****Bookmarks****








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
                    
            try
            {
                foreach (var v in Visuals.Where(a => a.PageName == Pages.Where(z => z.Id == rptPageId).FirstOrDefault().Name).ToList())
                {
                    string vizId = v.Id;
                    bool hm = false;
                    try
                    {  
                        string vT = (string)o["explorationState"]["sections"][rptPageId]["visualContainers"][vizId]["singleVisual"]["visualType"];    
                        
                        // visualContainers
                        try
                        {
                            string mode = (string)o["explorationState"]["sections"][rptPageId]["visualContainers"][vizId]["singleVisual"]["display"]["mode"];
                            if (mode == "hidden")
                            {
                                hm = true;
                            }
                        }
                        catch
                        {
                        }                                                
                        
                        Bookmarks.Add(new Bookmark {Id = bId, Name = bName, PageId = rptPageId, VisualId = vizId, VisualHiddenFlag = hm});
                    }
                    catch
                    {
                    }

                    // visualContainerGroups
                    try
                    {
                        bool mode = (bool)o["explorationState"]["sections"][rptPageId]["visualContainerGroups"][vizId]["isHidden"];
                        
                        if (mode)
                        {
                            hm = true;
                        }
                        
                        Bookmarks.Add(new Bookmark {Id = bId, Name = bName, PageId = rptPageId, VisualId = vizId, VisualHiddenFlag = hm});
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            } 
        }
    }
    catch
    {
    }

    // Update for Visual Groups
    foreach (var x in Visuals.ToList())
    {
        if (x.ParentGroup != null)
        {
            int pgX = Visuals.Where(a => a.Id == x.ParentGroup).First().X;
            int pgY = Visuals.Where(a => a.Id == x.ParentGroup).First().Y;
            
            x.X = x.X + pgX;
            x.Y = x.Y + pgY;
        }
    }    
    
    // Add results to StringBuilders
    foreach (var x in CustomVisuals.ToList())
    {
        sb_CustomVisuals.Append(ReportName + '\t' + x.Name + '\t' + ReportDate +  newline);
    }
    foreach (var x in ReportFilters.ToList())
    {
        sb_ReportFilters.Append(ReportName + '\t' + x.displayName + '\t' + x.TableName + '\t' + x.ObjectName + '\t' + x.ObjectType + '\t' + x.FilterType + '\t' + x.HiddenFilter + '\t' + x.LockedFilter + '\t' + x.AppliedFilterVersion + '\t' + ReportDate + newline);
    }
    foreach (var x in PageFilters.ToList())
    {
        sb_PageFilters.Append(ReportName + '\t' + x.PageId + '\t' + x.PageName + '\t' + x.displayName + '\t' + x.TableName + '\t' + x.ObjectName + '\t' + x.ObjectType + '\t' + x.FilterType + '\t' + x.HiddenFilter + '\t' + x.LockedFilter + '\t' + x.AppliedFilterVersion + '\t' + ReportDate + newline);
    }
    foreach (var x in VisualFilters.ToList())
    {
        sb_VisualFilters.Append(ReportName + '\t' + x.PageName + '\t' + x.VisualId + '\t' + x.TableName + '\t' + x.ObjectName + '\t' + x.ObjectType + '\t' + x.FilterType + '\t' + x.HiddenFilter + '\t' + x.LockedFilter + '\t' + x.AppliedFilterVersion + '\t' + x.displayName + '\t' + ReportDate + newline);
    }
    foreach (var x in VisualObjects.ToList())
    {
        sb_VisualObjects.Append(ReportName + '\t' + x.PageName + '\t' + x.VisualId + '\t' + x.VisualType + '\t' + x.AppliedFilterVersion + '\t' + x.CustomVisualFlag + '\t' + x.TableName + '\t' + x.ObjectName + '\t' + x.ObjectType + '\t' + x.Source + '\t' + x.displayName + '\t' + ReportDate + newline);
    }
    foreach (var x in Bookmarks.ToList())
    {
        sb_Bookmarks.Append(ReportName + '\t' + x.Name + '\t' + x.Id + '\t' + x.PageId + '\t' + x.VisualId + '\t' + x.VisualHiddenFlag + '\t' + ReportDate + newline);
    }
    foreach (var x in Pages.ToList())
    {
        sb_Pages.Append(ReportName + '\t' + x.Id + '\t' + x.Name + '\t' + x.Number + '\t' + x.Width + '\t' + x.Height + '\t' + x.HiddenFlag + '\t' + x.VisualCount + '\t' + x.BackgroundImage + '\t' + x.WallpaperImage + '\t' + x.Type + '\t' + ReportDate + newline);
    }
    foreach (var x in Visuals.ToList())
    {
        sb_Visuals.Append(ReportName + '\t' + x.PageName + '\t' + x.Id + '\t' + x.Name + '\t' + x.Type + '\t' + x.CustomVisualFlag + '\t' + x.HiddenFlag + '\t' + x.X + '\t' + x.Y + '\t' + x.Z + '\t' + x.Width + '\t' + x.Height + '\t' + x.ObjectCount + '\t' + x.ShowItemsNoDataFlag + '\t' + x.SlicerType + '\t' + x.ParentGroup + '\t' + ReportDate + newline);
    }
    foreach (var x in Connections.ToList())
    {
        sb_Connections.Append(ReportName + '\t' + x.ServerName + '\t' + x.DatabaseName + '\t' + x.ReportID + '\t' + x.Type + '\t' + ReportDate + newline);
    }
    foreach (var x in VisualInteractions.ToList())
    {
        sb_VisualInteractions.Append(ReportName + '\t' + x.PageName + '\t' + x.SourceVisualID + '\t' + x.TargetVisualID + '\t' + x.TypeID + '\t' + x.Type + '\t' + ReportDate +    newline);
    }
}

// Save to text files or print out results
if (saveToFile)
{    
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"CustomVisuals.txt", sb_CustomVisuals.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"ReportFilters.txt", sb_ReportFilters.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"PageFilters.txt", sb_PageFilters.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"VisualFilters.txt", sb_VisualFilters.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"VisualObjects.txt", sb_VisualObjects.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"Visuals.txt", sb_Visuals.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"Bookmarks.txt", sb_Bookmarks.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"Pages.txt", sb_Pages.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"Connections.txt", sb_Connections.ToString());
    System.IO.File.WriteAllText(pbiFolderName+@"\"+"VisualInteractions.txt", sb_VisualInteractions.ToString());
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
    sb_VisualInteractions.Output();
}

// Add dependencies to the perspective
if (createPersp)
{    
    // RLS
    var rlsColumnDependencies = 
    Model.Roles.SelectMany(r => r.TablePermissions)
        .SelectMany(tp => tp.DependsOn.Columns).Distinct().ToList();
        
    var rlsTableDependencies = Model.Roles.SelectMany(r => r.TablePermissions)
        .SelectMany(tp => tp.DependsOn.Tables).Distinct().ToList();

    var rlsMeasureDependencies = Model.Roles.SelectMany(r => r.TablePermissions)
        .SelectMany(tp => tp.DependsOn.Tables).Distinct().ToList();
        
    foreach (var x in rlsColumnDependencies)
    {
        x.InPerspective[perspName] = true;
    }
    foreach (var x in rlsTableDependencies)
    {
        x.InPerspective[perspName] = true;
    }
    foreach (var x in rlsMeasureDependencies)
    {
        x.InPerspective[perspName] = true;
    }

    //OLS
    foreach (var t in Model.Tables.ToList())
    {
        string tableName = t.Name;
        
        foreach(var r in Model.Roles.ToList())
        {
            string roleName = r.Name;
            string tableOLS = Model.Tables[tableName].ObjectLevelSecurity[roleName].ToString();
            if (tableOLS != "Default")
            {
                t.InPerspective[perspName] = true;
            }
            
            foreach (var c in t.Columns.ToList())
            {
                string colName = c.Name;
                string colOLS = Model.Tables[tableName].Columns[colName].ObjectLevelSecurity[roleName].ToString();
                
                if (colOLS != "Default")
                {
                    c.InPerspective[perspName] = true;
                }
            }
        }    
    }
    
    foreach (var o in Model.AllMeasures.Where(a => a.InPerspective[perspName]).ToList())
    {
        // Add measure dependencies
        var allReferences = o.DependsOn.Deep();
                            
        foreach(var dep in allReferences)
        {
            //Add dependent columns/measures specified in text file to the perspective
            var columnDep = dep as Column; if(columnDep != null) columnDep.InPerspective[perspName] = true;
            var measureDep = dep as Measure; if(measureDep != null) measureDep.InPerspective[perspName] = true;
        }
    }
    
    // Add hierarchy column dependencies
    foreach (var o in Model.AllHierarchies.Where(a => a.InPerspective[perspName]).ToList())
    {
        string tableName = o.Table.Name;
        foreach (var lev in o.Levels.ToList())
        {
            string hcolName = lev.Column.Name;
            Model.Tables[tableName].Columns[hcolName].InPerspective[perspName] = true;
        }
    }
    
    // Add foreign/primary keys
    foreach (var o in Model.Relationships.ToList())
    {
        var fromTable = o.FromTable;
        var toTable = o.ToTable;
        var fromColumn = o.FromColumn;
        var toColumn = o.ToColumn;
        
        if (fromTable.InPerspective[perspName] == true && toTable.InPerspective[perspName] == true)
        {
            fromColumn.InPerspective[perspName] = true;
            toColumn.InPerspective[perspName] = true;
        }
    }

    // Add auto-date tables
    if (Model.Tables.Any(a => a.DataCategory == "Time" && a.Columns.Any(b => b.IsKey && a.InPerspective[perspName])))
    {
        foreach (var t in Model.Tables.Where(a => a.ObjectTypeName == "Calculated Table" && (a.Name.StartsWith("DateTableTemplate_") || a.Name.StartsWith("LocalDateTable_"))).ToList())
        {
            t.InPerspective[perspName] = true;
        }
    }

    // Add all columns in calculation groups
    foreach (var t in Model.CalculationGroups.Where(a => a.InPerspective[perspName]).ToList())
    {
        foreach (var c in t.Columns.ToList())
        {
            c.InPerspective[perspName] = true;
        }
    }
    
    // Add column dependencies
    foreach (var c in Model.AllColumns.Where(a => a.InPerspective[perspName]).ToList())
    {
        // Add sort-by column dependencies
        try
        {
            c.SortByColumn.InPerspective[perspName] = true;
        }
        catch
        {
        }

        // Add calculated column dependencies
        if (c.Type.ToString() == "Calculated")
        {
            var allReferences = (Model.Tables[c.Table.Name].Columns[c.Name] as CalculatedColumn).DependsOn.Deep();

            foreach(var dep in allReferences)
            {
                // Add dependent columns/measures specified in text file to the perspective
                var columnDep = dep as Column; if(columnDep != null) columnDep.InPerspective[perspName] = true;
                var measureDep = dep as Measure; if(measureDep != null) measureDep.InPerspective[perspName] = true;
            }
        }
    }
}

// Show unused objects
if (createPersp)
{
    var sb_Unused = new System.Text.StringBuilder();
    sb_Unused.Append("TableName" + '\t' + "ObjectName" + '\t' + "ObjectType" + newline);

    foreach (var t in Model.Tables.ToList())
    {
        foreach (var o in t.Columns.Where(a => a.InPerspective[perspName] == false).ToList())
        {
            sb_Unused.Append(o.Table.Name + '\t' + o.Name + '\t' + "Column" + newline);
        }
        foreach (var o in t.Measures.Where(a => a.InPerspective[perspName] == false).ToList())
        {
            sb_Unused.Append(o.Table.Name + '\t' + o.Name + '\t' + "Measure" + newline);
        }
        foreach (var o in t.Hierarchies.Where(a => a.InPerspective[perspName] == false).ToList())
        {
            sb_Unused.Append(o.Table.Name + '\t' + o.Name + '\t' + "Hierarchy" + newline);
        }
    }

    if (saveToFile)
    {
        System.IO.File.WriteAllText(pbiFolderName+@"\"+savePrefix+"_UnusedObjects.txt", sb_Unused.ToString());
    }
    else
    {
        sb_Unused.Output();
    }
}

// Extra closing bracket for classes
} // Comment out this line if using Tabular Editor 3

// Classes for each object set
public class CustomVisual
{
    public string Name { get; set; }
    public string ReportDate { get; set; }
}

public class Bookmark
{
    public string Id { get; set; }
    public string Name { get; set; }
    public string PageId { get; set; }
    public string VisualId { get; set; }
    public bool VisualHiddenFlag { get; set; }
    public string ReportDate { get; set; }
}

public class ReportFilter
{
    public string displayName { get; set; }
    public string TableName { get; set; }
    public string ObjectName { get; set; }
    public string ObjectType { get; set; }
    public string FilterType { get; set; }
    public string HiddenFilter { get; set; }
    public string LockedFilter { get; set; }
    public string AppliedFilterVersion { get; set; }
    public string ReportDate { get; set; }
}

public class VisualObject
{
    public string PageName { get; set; }
    public string displayName { get; set; }
    public string VisualId { get; set; }
    public string VisualType { get; set; }
    public string AppliedFilterVersion { get; set; }
    public bool CustomVisualFlag { get; set; }
    public string TableName { get; set; }
    public string ObjectName { get; set; }
    public string ObjectType { get; set; }
    public string Source { get; set; }
    public string ReportDate { get; set; }
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
    public bool ShowItemsNoDataFlag { get; set; }
    public string SlicerType { get; set; }
    public string ParentGroup {get; set; }
    public string ReportDate { get; set; }
}

public class VisualFilter
{
    public string PageName { get; set; }
    public string VisualId { get; set; }
    public string TableName { get; set; }
    public string ObjectName { get; set; }
    public string ObjectType { get; set; }
    public string FilterType { get; set; }
    public string HiddenFilter { get; set; }
    public string LockedFilter { get; set; }
    public string AppliedFilterVersion { get; set; }
    public string displayName { get; set; }
    public string ReportDate { get; set; }
}

public class PageFilter
{
    public string PageId { get; set; }
    public string PageName { get; set; }
    public string displayName {get; set; }
    public string TableName { get; set; }
    public string ObjectName { get; set; }
    public string ObjectType { get; set; }
    public string FilterType { get; set; }   
    public string HiddenFilter { get; set; }
    public string LockedFilter { get; set; } 
    public string AppliedFilterVersion { get; set; }
    public string ReportDate { get; set; } 
}

public class Page
{
    public string Id { get; set; }
    public string Name { get; set; }
    public int Number { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public bool HiddenFlag { get; set; }
    public int VisualCount { get; set; }
    public string BackgroundImage { get; set; }
    public string WallpaperImage { get; set; }
    public string Type { get; set; }
    public string ReportDate { get; set; }
}

public class Connection
{
    public string ServerName { get; set; }
    public string DatabaseName { get; set; }
    public string Type { get; set; }
    public string ReportID { get; set; }
    public string ReportDate { get; set; }
}

public class VisualInteraction
{
    public string PageName { get; set; }
    public string SourceVisualID { get; set; }
    public string TargetVisualID { get; set; }
    public int TypeID { get; set; }
    public string Type { get; set; }
    public string ReportDate { get; set; }
}

static void _() { // Comment out this line if using Tabular Editor 3
