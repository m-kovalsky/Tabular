int maxLen = 100;
string annName = "LongLengthRowCount";

foreach (var c in Model.AllColumns.Where(a => a.DataType == DataType.String))
{
    string objName = c.DaxObjectFullName;
    var result = EvaluateDax("SUMMARIZECOLUMNS(\"test\",CALCULATE(COUNTROWS(DISTINCT("+objName+")),LEN("+objName+") > "+maxLen+"))");
    
    c.SetAnnotation(annName,result.ToString());
    
    if (c.GetAnnotation(annName) == "Table")
    {
        c.SetAnnotation(annName,"0");
    }
}