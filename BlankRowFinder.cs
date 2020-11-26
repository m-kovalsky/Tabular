var sb = new System.Text.StringBuilder();

sb.Append("FromTable" + '\t' + "ToTable" + '\t' + "BlankRowCount");
sb.Append(Environment.NewLine);

foreach (var r in Model.Relationships.ToList())
{
    bool   act = r.IsActive;
    string fromTable = r.FromTable.Name;
    string fromColumn = r.FromColumn.Name;
    string toTable = r.ToTable.Name;
    string toColumn = r.ToColumn.Name;
    string dax;
    
    if (act)
    {
        dax = "SUMMARIZECOLUMNS(\"test\",CALCULATE(COUNTROWS('"+fromTable+"'),ISBLANK('"+toTable+"'["+toColumn+"])))";
    }
    else
    {
        dax = "SUMMARIZECOLUMNS(\"test\",CALCULATE(COUNTROWS('"+fromTable+"'),USERELATIONSHIP('"+fromTable+"'["+fromColumn+"],'"+toTable+"'["+toColumn+"]),ISBLANK('"+toTable+"'["+toColumn+"])))";
    }
    
    var daxResult = EvaluateDax(dax);
    
    if (daxResult.ToString() != "Table")
    {
        sb.Append(fromTable + '\t' + toTable + '\t' + daxResult.ToString());
        sb.Append(Environment.NewLine);
    }
}

sb.Output();