#r "System.IO"
#r "Microsoft.Office.Interop.Excel"

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

string filePath = @"C:\Desktop\Descriptions"; // Update this to be the location of the Descriptions file
string excelFilePath = filePath + ".xlsx"; 
string excelTabName = "ModelDescriptions";

// Open Excel
var excelApp = new Excel.Application();
excelApp.Visible = false;
excelApp.DisplayAlerts = false;

// Open Workbook, Worksheet
var wb = excelApp.Workbooks.Open(excelFilePath); 
var ws = wb.Worksheets[excelTabName] as Excel.Worksheet;

// Count rows and columns
Excel.Range xlRange = ws.UsedRange;

int rowCount = xlRange.Rows.Count;

for (int r = 2; r <= rowCount; r++)
{
    string tableName = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
    string objType = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
    string objName = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString();
    string desc = (string)(ws.Cells[r,5] as Excel.Range).Text.ToString();
    
    if (objType == "Table")
    {
        try
        {
            Model.Tables[tableName].Description = desc;
        }
        catch
        {
        }
    }
    else if (objType == "Column")
    {
        try
        {
            Model.Tables[tableName].Columns[objName].Description = desc;
        }
        catch
        {            
        }
    }
    else if (objType == "Measure")
    {
        try
        {
            Model.Tables[tableName].Measures[objName].Description = desc;
        }
        catch
        {
        }
    }
    else if (objType == "Hierarchy")
    {
        try
        {
            Model.Tables[tableName].Hierarchies[objName].Description = desc;
        }
        catch
        {
        }
    }
    else if (objType == "Calculation Group")
    {
        try
        {
            Model.Tables[tableName].Description = desc;
        }
        catch
        {
        }
    }
    else if (objType == "Calculation Item")
    {
        try
        {
            (Model.Tables[tableName] as CalculationGroupTable).CalculationItems[objName].Description = desc;
        }
        catch
        {
        }
    }
}

// Close workbook and quit Excel program
wb.Close();
excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);