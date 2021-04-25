var sb = new System.Text.StringBuilder(); 
string newline = Environment.NewLine;
string ann = "Vertipaq_ColumnSize";
long tot = 0;

// Header
sb.Append("TableName" + '\t' + "ColumnName" + '\t' + "ColumnSize" + newline );

foreach (var c in Model.AllColumns.Where(a => (a.IsHidden || a.Table.IsHidden) && a.ReferencedBy.Count() == 0 && ! a.UsedInRelationships.Any() && ! a.UsedInSortBy.Any() && ! a.UsedInHierarchies.Any() 
    && (! a.Table.RowLevelSecurity.Any(b => b != null && b.IndexOf("[" + a.Name + "]", StringComparison.OrdinalIgnoreCase) >= 0 ))
    && (! Model.Roles.Any(c => c.RowLevelSecurity.Any(d => d != null && (d.IndexOf(a.Table.Name + "[" + a.Name + "]", StringComparison.OrdinalIgnoreCase) >=0 || d.IndexOf("'" + a.Table.Name + "'[" + a.Name + "]", StringComparison.OrdinalIgnoreCase) >=0))))).OrderBy(a => a.Table.Name).ThenBy(a => a.Name))
    
{
    string tableName = c.Table.Name;
    string colName = c.Name;
    string annValue = c.GetAnnotation(ann);
    sb.Append(tableName + '\t' + colName + '\t' + annValue + newline);
    tot = tot + Convert.ToInt64(annValue);
}

tot.Output(); // Value shown in bytes
sb.ToString().Output();