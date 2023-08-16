var sb = new System.Text.StringBuilder();
string newline = Environment.NewLine;

sb.Append("TableName" + '\t' + "MeasureName");

// Fix table names in measures
foreach (var m in Model.AllMeasures.ToList())
{
    string origExpr = m.Expression;
    string expr = m.Expression.ToLower();
    string tableMeasure = m.Table.Name + '\t' + m.Name;
    
    foreach (var t in Model.Tables.Where(a => expr.Contains(a.Name.ToLower() + "[") || expr.Contains(a.Name.ToLower() + "'[")))
    {        
        m.Expression = m.Expression.Replace(t.Name,t.Name,StringComparison.OrdinalIgnoreCase);
        if (origExpr != m.Expression && !sb.ToString().Contains(tableMeasure))
        {
            sb.Append(newline + tableMeasure);
        }
    }
}

// Fix column names in measures
foreach (var m in Model.AllMeasures.ToList())
{
    string origExpr = m.Expression;
    string expr = m.Expression.ToLower();
    string tableMeasure = m.Table.Name + '\t' + m.Name;
 
    foreach (var c in Model.AllColumns.Where(a => expr.Contains(a.Table.Name.ToLower() + "[" + a.Name.ToLower() + "]") || expr.Contains("'" + a.Table.Name.ToLower() + "'[" + a.Name.ToLower() + "]")))
    {
        m.Expression = m.Expression.Replace(c.Name,c.Name,StringComparison.OrdinalIgnoreCase);
        if (origExpr != m.Expression && !sb.ToString().Contains(tableMeasure))
        {
            sb.Append(newline + tableMeasure);
        }
    }   
}

// Fix measure names in measures
foreach (var m in Model.AllMeasures.ToList())
{
    string origExpr = m.Expression;
    string expr = m.Expression.ToLower();
    string tableMeasure = m.Table.Name + '\t' + m.Name;
 
    foreach (var c in Model.AllMeasures.Where(a => expr.Contains("[" + a.Name.ToLower() + "]")))
    {
        m.Expression = m.Expression.Replace(c.Name,c.Name,StringComparison.OrdinalIgnoreCase);
        if (origExpr != m.Expression && !sb.ToString().Contains(tableMeasure))
        {
            sb.Append(newline + tableMeasure);
        }
    }   
}

Info("Object references have been updated to match the actual table, column or measure name within the model for the following measures:" + newline + newline + sb.ToString());
