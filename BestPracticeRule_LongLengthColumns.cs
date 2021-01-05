int maxLen = 100;
string annName = "LongLengthRowCount";

foreach (var c in Model.AllColumns.Where(a => a.DataType == DataType.String))
{
    string tableName = c.Table.Name;
    string columnName = c.Name;
    
    var obj = Model.Tables[tableName].Columns[columnName];
    var result = EvaluateDax("SUMMARIZECOLUMNS(\"test\",CALCULATE(COUNTROWS(DISTINCT('"+tableName+"'["+columnName+"])),LEN('"+tableName+"'["+columnName+"]) > "+maxLen+"))");
    
    obj.SetAnnotation(annName,result.ToString());
    
    if (obj.GetAnnotation(annName) == "Table")
    {
        obj.SetAnnotation(annName,"0");
    }
}