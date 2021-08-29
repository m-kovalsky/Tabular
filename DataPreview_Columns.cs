string dax = string.Empty;

foreach (var x in Selected.Columns.ToList())
{
    dax = dax + x.DaxObjectFullName + ",";
}

dax = dax.Substring(0,dax.Length-1);
var result = EvaluateDax("TOPN(500,SUMMARIZECOLUMNS("+dax+"))");
result.Output();