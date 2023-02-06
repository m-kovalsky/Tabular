var sb = new System.Text.StringBuilder();
string newline = Environment.NewLine;

sb.Append("TableName" + '\t' + "MeasureName");

foreach (var m in Model.AllMeasures.ToList())
{
    var allReferences = m.DependsOn.Deep();
    
    if (!allReferences.Any(a => a.ObjectType.ToString() == "Measure"))
    {
        sb.Append(newline + m.Table.Name + '\t' + m.Name);
    }
}

sb.Output();