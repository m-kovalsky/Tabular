var colNames = Selected.Columns.ToList();
string dax = string.Empty;

for (int i=0; i<colNames.Count(); i++)
{
    string colFullName = colNames[i].DaxObjectFullName;
    dax = dax + colFullName + ",";
}

dax = dax.Substring(0,dax.Length-1);
var result = EvaluateDax("TOPN(500,SUMMARIZECOLUMNS("+dax+"))");
result.Output();