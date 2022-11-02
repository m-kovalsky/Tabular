foreach (var col in Model.AllColumns.Where(c=>c.Name.Contains("GUID_")))
{
    col.IsHidden=true;
}