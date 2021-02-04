#r "System.IO"
#r "System.IO.Compression.FileSystem"

using System.IO;
using System.IO.Compression;

string vpaxFile = @"C:\Desktop\ModelVertipaq.vpax"; // Enter .vpax file path

string fileExt = Path.GetExtension(vpaxFile);

if (fileExt != ".vpax")
{
    Error("Must use a valid .vpax file");
}

string fileName = Path.GetFileNameWithoutExtension(vpaxFile);
string folderName = Path.GetDirectoryName(vpaxFile) + @"\";
string zipPath = folderName + fileName + ".zip";
string unzipPath = folderName + fileName;

try
{
    // Make a copy of a vpax and turn it into a zip file
    File.Copy(vpaxFile, zipPath);
    // Unzip file
    System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, unzipPath);
    // Delete zip file
    File.Delete(zipPath);
}

catch
{
    Error("File does not exist. Must use a valid .vpax file");
}

// Remove Existing Vertipaq Annotations
Model.RemoveAnnotation("Vertipaq_ModelSize");

foreach (var o in Model.AllHierarchies)
{
    o.RemoveAnnotation("Vertipaq_UserHierarchySize");
    o.RemoveAnnotation("Vertipaq_TableSizePctOfModel");
}

foreach (var o in Model.AllColumns)
{
    o.RemoveAnnotation("Vertipaq_ColumnHierarchySize");
    o.RemoveAnnotation("Vertipaq_DataSize");
    o.RemoveAnnotation("Vertipaq_DictionarySize");
    o.RemoveAnnotation("Vertipaq_Cardinality");
    o.RemoveAnnotation("Vertipaq_ColumnSize");
    o.RemoveAnnotation("Vertipaq_ColumnSizePctOfTable");
    o.RemoveAnnotation("Vertipaq_ColumnSizePctOfModel");
}

foreach (var o in Model.Relationships.ToList())
{        
    o.RemoveAnnotation("Vertipaq_RelationshipSize");   
    o.RemoveAnnotation("Vertipaq_MaxFromCardinality");   
    o.RemoveAnnotation("Vertipaq_MaxToCardinality");        
}

foreach (var o in Model.Tables.ToList())
{     
    o.RemoveAnnotation("Vertipaq_RowCount");
    o.RemoveAnnotation("Vertipaq_TableSize");
}

foreach (var o in Model.AllPartitions)
{       
    o.RemoveAnnotation("Vertipaq_RecordCount");
    o.RemoveAnnotation("Vertipaq_RecordsPerSegment");
    o.RemoveAnnotation("Vertipaq_SegmentCount");
}

// Deseralize json file
string jsonFilePath = folderName + fileName + @"\" + "DaxVpaView.json";
var unformattedJson = File.ReadAllText(jsonFilePath,System.Text.UnicodeEncoding.Unicode);
var formattedJson = Newtonsoft.Json.Linq.JToken.Parse(unformattedJson).ToString();

dynamic json = Newtonsoft.Json.Linq.JObject.Parse(formattedJson);

// Delete previously created folder
try
{
    Directory.Delete(folderName + fileName,true);
}
catch
{
}

int tableCount = (int)json["Tables"].Count;
int columnCount = (int)json["Columns"].Count;
int relationshipCount = (int)json["Relationships"].Count;
int hierarchiesCount = (int)json["UserHierarchies"].Count;
int columnSegmentCount = (int)json["ColumnsSegments"].Count;

// Add table annotations
for (int i=0; i<tableCount; i++)
{
    string tableName = (string)json["Tables"][i]["TableName"];
    string rowCount = (string)json["Tables"][i]["RowsCount"];
    string tableSize = (string)json["Tables"][i]["TableSize"];
    
    if (Model.Tables.Where(a => a.Name == tableName).Count() == 1)
    {
        var obj = Model.Tables[tableName];
        
        obj.SetAnnotation("Vertipaq_RowCount",rowCount);
        obj.SetAnnotation("Vertipaq_TableSize",tableSize);
    }
}

// Add column annotations
for (int i=0; i<columnCount; i++)
{
    string columnName = (string)json["Columns"][i]["ColumnName"];
    string tableName = (string)json["Columns"][i]["TableName"];
    string columnCardinality = (string)json["Columns"][i]["ColumnCardinality"];
    string dictionarySize = (string)json["Columns"][i]["DictionarySize"];
    string dataSize = (string)json["Columns"][i]["DataSize"];
    string hierarchiesSize = (string)json["Columns"][i]["HierarchiesSize"];
    string totalSize = (string)json["Columns"][i]["TotalSize"];
    
    if (Model.Tables.Where(a => a.Name == tableName && a.Columns.Any(b => b.Name == columnName)).Count() == 1)
    
    {
        var obj = Model.Tables[tableName].Columns[columnName];
    
        obj.SetAnnotation("Vertipaq_Cardinality",columnCardinality);
        obj.SetAnnotation("Vertipaq_ColumnHierarchySize",hierarchiesSize);
        obj.SetAnnotation("Vertipaq_ColumnSize",totalSize);
        obj.SetAnnotation("Vertipaq_DataSize",dataSize);
        obj.SetAnnotation("Vertipaq_DictionarySize",dictionarySize);
    }      
}

// Add relationship annotations
for (int i=0; i<relationshipCount; i++)
{
    string relationshipName = (string)json["Relationships"][i]["RelationshipName"];
    string fromCardinality = (string)json["Relationships"][i]["FromCardinality"];
    string toCardinality = (string)json["Relationships"][i]["ToCardinality"];
    string usedSize = (string)json["Relationships"][i]["UsedSize"];
    
    if (Model.Relationships.Where(a => a.ID == relationshipName).Count() == 1)
    {
        var obj = Model.Relationships[relationshipName];
    
        obj.SetAnnotation("Vertipaq_MaxFromCardinality",fromCardinality);
        obj.SetAnnotation("Vertipaq_MaxToCardinality",toCardinality);
        obj.SetAnnotation("Vertipaq_RelationshipSize",usedSize);
    }
}

// Add hierarchies annotations
for (int i=0; i<hierarchiesCount; i++)
{
    string hierarchyName = (string)json["UserHierarchies"][i]["UserHierarchyName"];
    string tableName = (string)json["UserHierarchies"][i]["TableName"];
    string usedSize = (string)json["UserHierarchies"][i]["UsedSize"];
    
    if (Model.AllHierarchies.Where(a => a.Name == hierarchyName && a.Table.Name == tableName).Count() == 1)
    {
        var obj = Model.Tables[tableName].Hierarchies[hierarchyName];
    
        obj.SetAnnotation("Vertipaq_UserHierarchySize",usedSize);
    }
}

// Add partition annotations
for (int i=0; i<columnSegmentCount; i++)
{
    string tableName = (string)json["ColumnsSegments"][i]["TableName"];
    string partitionName = (string)json["ColumnsSegments"][i]["PartitionName"];
    string columnName = (string)json["ColumnsSegments"][i]["ColumnName"];
    string segmentNumber = (string)json["ColumnsSegments"][i]["SegmentNumber"];
    string tablePartitionNumber = (string)json["ColumnsSegments"][i]["TablePartitionNumber"];
    string segmentRows = (string)json["ColumnsSegments"][i]["SegmentRows"];
    int segmentNumberInt = Convert.ToInt32(segmentNumber);
    int tablePartitionNumberInt = Convert.ToInt32(tablePartitionNumber);
    long segmentRowsInt = Convert.ToInt64(segmentRows);
    
    var obj = Model.Tables[tableName].Partitions[partitionName];
    
    int s = 0;
    foreach (var t in Model.Tables.Where(a => a.Name == tableName).ToList())
    {
        foreach (var p in t.Partitions.Where(b => b.MetadataIndex < tablePartitionNumberInt))
        {
            s = s + Convert.ToInt32(p.GetAnnotation("Vertipaq_SegmentCount"));
        }
    }
    
    obj.SetAnnotation("Vertipaq_SegmentCount",(segmentNumberInt - s + 1).ToString());
    
    if (columnName.StartsWith("RowNumber-"))
    {            
        long rc = Convert.ToInt64(obj.GetAnnotation("Vertipaq_RecordCount"));
        obj.SetAnnotation("Vertipaq_RecordCount",(segmentRowsInt + rc).ToString());
    }
}

// Add Records per Segment
long maxRPS = 8388608;
foreach (var t in Model.Tables.ToList())
{
    foreach (var p in t.Partitions.ToList())
    {
        long rc = Convert.ToInt64(p.GetAnnotation("Vertipaq_RecordCount"));
        long sc = Convert.ToInt64(p.GetAnnotation("Vertipaq_SegmentCount"));
        string rps = "Vertipaq_RecordsPerSegment";
        
        if (sc > 1)
        {
            p.SetAnnotation(rps,maxRPS.ToString());            
        }
        else if (sc == null || sc == 0)
        {
            p.SetAnnotation(rps,"0");
        }
        else
        {
            p.SetAnnotation(rps,(rc / sc).ToString());
        }
    }
}

// Add model size annotation
string ms = Model.Tables.Sum(a => Convert.ToInt64(a.GetAnnotation("Vertipaq_TableSize"))).ToString();
Model.SetAnnotation("Vertipaq_ModelSize",ms);

// Percent of Table and Model
float modelSize = Convert.ToInt64(Model.GetAnnotation("Vertipaq_ModelSize"));

foreach (var t in Model.Tables.ToList())
{
    string tableName = t.Name;
    var obj = Model.Tables[tableName];
    
    float tableSize = Convert.ToInt64(obj.GetAnnotation("Vertipaq_TableSize"));
    double tblpct = Math.Round(tableSize / modelSize,3);
        
    obj.SetAnnotation("Vertipaq_TableSizePctOfModel",tblpct.ToString());
    
    foreach (var c in t.Columns.ToList())
    {
        string colName = c.Name;
        var col = Model.Tables[tableName].Columns[colName];
        
        float colSize = Convert.ToInt64(col.GetAnnotation("Vertipaq_ColumnSize"));
        double colpctTbl = Math.Round(colSize / tableSize,3);
        double colpctModel = Math.Round(colSize / modelSize,3);
        
        col.SetAnnotation("Vertipaq_ColumnSizePctOfTable",colpctTbl.ToString());
        col.SetAnnotation("Vertipaq_ColumnSizePctOfModel",colpctModel.ToString());
    }
}