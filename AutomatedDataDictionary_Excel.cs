#r "System.IO"
#r "Microsoft.Office.Interop.Excel"

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

string filePath = @"C:\Desktop\DataDictionary"; // Update this to be the desired location of the Data Dictionary file
bool dataSourceM = false; // Set this to true if you want the data source to use M
string excelFilePath = filePath + ".xlsx"; 
string textFilePath = filePath + ".txt";
string modelName = Model.Database.Name;
string ddName = "Data Dictionary";
string ddSource = "Excel " + ddName;
string conn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+excelFilePath+";Persist Security Info=false;Extended Properties=\"Excel 12.0;HDR=Yes;\"";
string[] colName = { "Model","Table","Object Type","Object","Hidden Flag","Description","Display Folder","Measure Formula" };
int colNameCount = colName.Length;
var sb = new System.Text.StringBuilder();
string newline = Environment.NewLine;

if (modelName == "SemanticModel")
{
    Error("Please name your model in the properties window: Model -> Database -> Name");
    return;
}

// Create Structured Data Source (M-Partitions)
if (dataSourceM)
{
    // Add data source
    if (!Model.DataSources.Any(a => a.Name == ddSource))
    {
        var ds = Model.AddStructuredDataSource(ddSource);
        ds.Protocol = "file";
        ds.Path = excelFilePath;
        ds.AuthenticationKind = "ServiceAccount";
        ds.ContextExpression = "let" + newline + "#\"0001\" = Excel.Workbook(..., null, true)" + newline + "in" + newline + "#\"0001\"";
    }
    
    // Add table
    if (!Model.Tables.Any(a => a.Name == ddName))
    {
        var t = Model.AddTable(ddName);
        t.AddMPartition(ddName+"1");
        t.Partitions[0].Delete();
        var p = (Model.Tables[ddName].Partitions[0] as MPartition);
        p.Name = ddName;
        p.MExpression = "let" + newline + 
        "    Source = #\""+ddSource+"\"," + newline + 
        "    #\""+ddName+"_Sheet\" = Source{[Item=\""+ddName+"\",Kind=\"Sheet\"]}[Data]," + newline +
        "    #\"Changed Type\" = Table.TransformColumnTypes(#\""+ddName+"_Sheet\",{{\"Column1\", type text}, {\"Column2\", type text}, {\"Column3\", type text}, {\"Column4\", type text}, {\"Column5\", type text}, {\"Column6\", type text}, {\"Column7\", type text}, {\"Column8\", type text}})," + newline +
        "    #\"Promoted Headers\" = Table.PromoteHeaders(#\"Changed Type\", [PromoteAllScalars=true])," + newline +
        "    #\"Changed Type1\" = Table.TransformColumnTypes(#\"Promoted Headers\",{{\"Model\", type text}, {\"Table\", type text}, {\"Object Type\", type text}, {\"Object\", type text}, {\"Hidden Flag\", type text}, {\"Description\", type text}, {\"Display Folder\", type text}, {\"Measure Formula\", type text}})" + newline +
        "in" + newline +
        "    #\"Changed Type1\"";
        
                
        // Add columns
        for (int i=0; i<colNameCount; i++)
        {
            var col = t.AddDataColumn(colName[i]);
            col.SourceColumn = colName[i];
            col.DataType = DataType.String;
        }
    }
}

// Create Legacy Data Source
else
{
    // Add data source
    if (!Model.DataSources.Any(a => a.Name == ddSource))
    {
        var ds = Model.AddDataSource(ddSource);
        ds.ConnectionString = conn;
    }
    
    // Add table
    if (!Model.Tables.Any(a => a.Name == ddName))
    {
        var t = Model.AddTable(ddName);
        t.Partitions[0].DataSource = Model.DataSources[ddSource];
        t.Partitions[0].Query = "SELECT * FROM ["+ddName+"$]";
        
        // Add columns
        for (int i=0; i<colNameCount; i++)
        {
            var col = t.AddDataColumn(colName[i]);
            col.SourceColumn = colName[i];
            col.DataType = DataType.String;
        }
    }
}

// Add headers
for (int i=0; i < colNameCount; i++)
{
    if (i<colNameCount-1)
    {
        sb.Append(colName[i] + '\t');
    }
    else
    {
        sb.Append(colName[i] + newline);
    }
}

// Extract model metadata in the data dictionary format
foreach (var t in Model.Tables.Where(a => a.ObjectType.ToString() != "CalculationGroupTable" && a.Name != ddName).OrderBy(a => a.Name).ToList())
{
    string tableName = t.Name;
    string tableDesc = t.Description.Replace("'","''");
    string objectType = "Table";
    string hiddenFlag;                 
    string expr;

    if (t.IsHidden)
    {
        hiddenFlag = "Yes";
    }
    else
    {
        hiddenFlag = "No";
    }
    
    if (t.SourceType.ToString() == "Calculated")
    {
        expr = (Model.Tables[tableName] as CalculatedTable).Expression;
 
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ").Replace("\t"," ");

        sb.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + tableName + '\t' + hiddenFlag + '\t' + tableDesc + '\t' + " " + '\t' + expr + newline);
    }
    else
    {
        sb.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + tableName + '\t' + hiddenFlag + '\t' + tableDesc + '\t' + " " + '\t' + "***N/A***" + newline);
    }
    
    foreach (var o in t.Columns.OrderBy(a => a.Name).ToList())
    {
        string objectName = o.Name;
        string objectDesc = o.Description;
        string objectDF = o.DisplayFolder;
        objectType = "Attribute";
        
        if (o.IsHidden)
        {
            hiddenFlag = "Yes";
        }
        else
        {
            hiddenFlag = "No";
        }
        
        if (o.Type.ToString() == "Calculated")
        {
            expr = (Model.Tables[tableName].Columns[objectName] as CalculatedColumn).Expression;
            
            // Remove tabs and new lines
            expr = expr.Replace("\n"," ").Replace("\t"," ");

            sb.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + objectDF + '\t' + expr + newline);        
        }
        else
        {
            sb.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + objectDF + '\t' + "***N/A***" + newline); 
        }
    }
    
    foreach (var o in t.Measures.OrderBy(a => a.Name).ToList())
    {
        string objectName = o.Name;
        string objectDesc = o.Description;
        string objectDF = o.DisplayFolder;
        objectType = "Measure";
        expr = o.Expression;                    
        
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ").Replace("\t"," ");
        
        if (o.IsHidden)
        {
            hiddenFlag = "Yes";
        }
        else
        {
            hiddenFlag = "No";
        }
        
        sb.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + objectDF + '\t' + expr + newline);
    }
    
    foreach (var o in t.Hierarchies.OrderBy(a => a.Name).ToList())
    {
        string objectName = o.Name;
        string objectDesc = o.Description;
        string objectDF = o.DisplayFolder;
        objectType = "Hierarchy";
        
        if (o.IsHidden)
        {
            hiddenFlag = "Yes";
        }
        else
        {
            hiddenFlag = "No";
        }
        
        sb.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + objectDF + '\t' + "***N/A***" + newline);
    }
}

foreach (var o in Model.CalculationGroups.ToList())
{
    string tableName = o.Name;
    string tableDesc = o.Description;
    string hiddenFlag;
    string objectType = "Calculation Group";
    
    if (o.IsHidden)
    {
        hiddenFlag = "Yes";
    }
    else
    {
        hiddenFlag = "No";
    }
    
    sb.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + tableName + '\t' + hiddenFlag + '\t' + tableDesc + '\t' + "" + '\t' + "***N/A***" + newline);    
    
    foreach (var i in o.CalculationItems.ToList())
    {        
        string objectName = i.Name;
        string objectDesc = i.Description;        
        string expr = i.Expression;
        objectType = "Calculation Item";
        
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ");
        expr = expr.Replace("\t"," ");
        expr = expr.Replace("'","''");
        
        sb.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + "" + '\t' + expr + newline);    
    }
} 

// Delete existing text/Excel files
try
{
    File.Delete(textFilePath);
    File.Delete(excelFilePath);
}
catch
{
}

// Save to text file
SaveFile(textFilePath, sb.ToString());

// Save to Excel file
var excelApp = new Excel.Application();
excelApp.Visible = false;
excelApp.DisplayAlerts = false;
excelApp.Workbooks.OpenText(textFilePath, Type.Missing, 1, Excel.XlTextParsingType.xlDelimited, Excel.XlTextQualifier.xlTextQualifierNone, false, true, false, false, false, false, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing);

var wb = excelApp.ActiveWorkbook;
var ws = wb.ActiveSheet as Excel.Worksheet;
ws.Name = ddName;
wb.SaveAs(excelFilePath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange);

// Close workbook and quit Excel program
wb.Close();
excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

// Delete text file as it is no longer necessary
try
{
    File.Delete(textFilePath);
}
catch
{
}

