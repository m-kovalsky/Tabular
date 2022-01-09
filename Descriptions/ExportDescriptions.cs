#r "System.IO"
#r "Microsoft.Office.Interop.Excel"

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

string filePath = @"C:\Desktop\Descriptions"; // Update this to be the desired location of the Descriptions file
string excelFilePath = filePath + ".xlsx"; 
string textFilePath = filePath + ".txt";
string excelTabName = "ModelDescriptions";
var sb = new System.Text.StringBuilder();
string newline = Environment.NewLine;

sb.Append("TableName" + '\t' + "ObjectType" + '\t' + "ObjectName" + '\t' + "HiddenFlag" + '\t' + "Description" + newline);

foreach (var t in Model.Tables.Where(a => a.ObjectType.ToString() != "CalculationGroupTable").OrderBy(a => a.Name).ToList())
{
    string tableName = t.Name;
    string tableDesc = t.Description;
    string tblhid;
    
    if (t.IsHidden)
    {
        tblhid = "Yes";
    }
    else
    {
        tblhid = "No";
    }
    
    sb.Append(tableName + '\t' + "Table" + '\t' + tableName + '\t' + tblhid + '\t' + tableDesc + newline);
    
    foreach (var o in t.Columns.OrderBy(a => a.Name).ToList())
    {
        string objName = o.Name;
        string objDesc = o.Description;
        string objhid;
    
        if (o.IsHidden)
        {
            objhid = "Yes";
        }
        else
        {
            objhid = "No";
        }
        
        sb.Append(tableName + '\t' + "Column" + '\t' + objName + '\t' + objhid + '\t' + objDesc + newline);        
    }
    
    foreach (var o in t.Measures.OrderBy(a => a.Name).ToList())
    {
        string objName = o.Name;
        string objDesc = o.Description;
        string objhid;
    
        if (o.IsHidden)
        {
            objhid = "Yes";
        }
        else
        {
            objhid = "No";
        }
        
        sb.Append(tableName + '\t' + "Measure" + '\t' + objName + '\t' + objhid + '\t' + objDesc + newline);        
    }
    
    foreach (var o in t.Hierarchies.OrderBy(a => a.Name).ToList())
    {
        string objName = o.Name;
        string objDesc = o.Description;
        string objhid;
    
        if (o.IsHidden)
        {
            objhid = "Yes";
        }
        else
        {
            objhid = "No";
        }
        
        sb.Append(tableName + '\t' + "Hierarchy" + '\t' + objName + '\t' + objhid + '\t' + objDesc + newline);        
    }    
}

foreach (var o in Model.CalculationGroups.OrderBy(a => a.Name).ToList())
{
    string tableName = o.Name;
    string tableDesc = o.Description;
    string tblhid;
    
    if (o.IsHidden)
    {
        tblhid = "Yes";
    }
    else
    {
        tblhid = "No";
    }
    
    sb.Append(tableName + '\t' + "Calculation Group" + '\t' + tableName + '\t' + tblhid + '\t' + tableDesc + newline);  
    
    foreach (var i in o.CalculationItems.ToList())
    {        
        string objName = i.Name;
        string objDesc = i.Description;
        
        sb.Append(tableName + '\t' + "Calculation Item" + '\t' + objName + '\t' + "No" + '\t' + objDesc + newline);        
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
excelApp.Workbooks.OpenText(textFilePath, 65001, 1, Excel.XlTextParsingType.xlDelimited, Excel.XlTextQualifier.xlTextQualifierNone, false, true, false, false, false, false, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing);

var wb = excelApp.ActiveWorkbook;
var ws = wb.ActiveSheet as Excel.Worksheet;
ws.Name = excelTabName;
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