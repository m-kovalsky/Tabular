string tableName = Selected.Table.Name;

var result = EvaluateDax("TOPN(10,'"+tableName+"')");
result.Output();