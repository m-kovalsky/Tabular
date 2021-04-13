var sb = new System.Text.StringBuilder();
string newline = Environment.NewLine;

sb.Append("FromTable" + '\t' + "ToTable" + '\t' + "BlankRowCount" + newline);

foreach (var r in Model.Relationships.ToList())
{
    bool   act = r.IsActive;
    string fromTable = r.FromTable.Name;
    string toTable = r.ToTable.Name;
    string fromTableFull = r.FromTable.DaxObjectFullName;    
    string fromObject = r.FromColumn.DaxObjectFullName;
    string toObject = r.ToColumn.DaxObjectFullName;
    string dax;
    
    if (act)
    {
        dax = "SUMMARIZECOLUMNS(\"test\",CALCULATE(COUNTROWS("+fromTableFull+"),ISBLANK("+toObject+")))";
    }
    else
    {
        dax = "SUMMARIZECOLUMNS(\"test\",CALCULATE(COUNTROWS("+fromTableFull+"),USERELATIONSHIP("+fromObject+","+toObject+"),ISBLANK("+toObject+")))";
    }
    
    var daxResult = EvaluateDax(dax);
    string blankRowCount = daxResult.ToString();
    
    if (blankRowCount != "Table")
    {
        sb.Append(fromTable + '\t' + toTable + '\t' + blankRowCount + newline);        
    }
}

sb.Output();