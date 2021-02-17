var sb = new System.Text.StringBuilder(); 

// Headers
sb.Append("TableName" + '\t' + "ColumnName" + '\t' + "HierarchySize" );
sb.Append(Environment.NewLine);

string ann = "Vertipaq_ColumnHierarchySize";
long tot = 0;
foreach (var c in Model.AllColumns.Where(a => a.IsAvailableInMDX && (a.IsHidden || a.Table.IsHidden) && ! a.UsedInSortBy.Any() && ! a.UsedInHierarchies.Any() ).OrderBy(a => a.Table.Name).ThenBy(a => a.Name))
{
    sb.Append(c.Table.Name + '\t' + c.Name + '\t' + c.GetAnnotation(ann));
    tot = tot + Convert.ToInt64(c.GetAnnotation(ann));
    sb.Append(Environment.NewLine);
}

tot.Output(); // Value shown in bytes
sb.ToString().Output();