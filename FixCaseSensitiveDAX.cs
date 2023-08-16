// Fix table names in measures
foreach (var m in Model.AllMeasures.ToList())
{
    string expr = m.Expression.ToLower();
    
    foreach (var t in Model.Tables.Where(a => expr.Contains(a.Name.ToLower() + "[") || expr.Contains(a.Name.ToLower() + "'[")))
    {
        m.Expression = m.Expression.Replace(t.Name,t.Name,StringComparison.OrdinalIgnoreCase);
    }
}

//Fix column names in measures
foreach (var m in Model.AllMeasures.ToList())
{
    string expr = m.Expression.ToLower();
 
    foreach (var c in Model.AllColumns.Where(a => expr.Contains(a.Table.Name.ToLower() + "[" + a.Name.ToLower() + "]") || expr.Contains("'" + a.Table.Name.ToLower() + "'[" + a.Name.ToLower() + "]")))
    {
        m.Expression = m.Expression.Replace(c.Name,c.Name,StringComparison.OrdinalIgnoreCase);
    }   
}

// Fix measure names in measures
foreach (var m in Model.AllMeasures.ToList())
{
    string expr = m.Expression.ToLower();
 
    foreach (var c in Model.AllMeasures.Where(a => expr.Contains("[" + a.Name.ToLower() + "]")))
    {
        m.Expression = m.Expression.Replace(c.Name,c.Name,StringComparison.OrdinalIgnoreCase);
    }   
}
