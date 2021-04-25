var sb = new System.Text.StringBuilder(); 
string newline = Environment.NewLine;
string ann = "Vertipaq_ColumnHierarchySize";
long tot = 0;

// Header
sb.Append("TableName" + '\t' + "ColumnName" + '\t' + "HierarchySize" + newline);

foreach (var c in Model.AllColumns.Where(a => a.IsAvailableInMDX && (a.IsHidden || a.Table.IsHidden) && ! a.UsedInSortBy.Any() && ! a.UsedInHierarchies.Any() ).OrderBy(a => a.Table.Name).ThenBy(a => a.Name))
{
    string tableName = c.Table.Name;
    string colName = c.Name;
    string annValue = c.GetAnnotation(ann);
    sb.Append(tableName + '\t' + colName + '\t' + annValue + newline);
    tot = tot + Convert.ToInt64(annValue);
}

tot.Output(); // Value shown in bytes
sb.ToString().Output();