List<string> TableList = new List<string>(); 

int i=0;
foreach (var t in Selected.Tables.ToList())
{
    if (i==0)
    {
        foreach (var x in t.RelatedTables)
        {
            TableList.Add(x.Name);
        }
    }
    else
    {
        foreach (var x in t.RelatedTables)
        {
            if (!TableList.Any(a => a == x.Name))
            {
                TableList.Remove(x.Name);
            }
        }
    }
    i++;
}

TableList.OrderBy(a => a).Output();