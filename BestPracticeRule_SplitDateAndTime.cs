// Split Date/Time recommendation
string annName = "DateTimeWithHourMinSec";
foreach (var c in Model.AllColumns.Where(a => a.DataType == DataType.DateTime))
{
    string columnName = c.Name;
    string tableName = c.Table.Name;
    var obj = Model.Tables[tableName].Columns[columnName];
    
    var result = ExecuteDax("EVALUATE TOPN(5,SUMMARIZECOLUMNS('"+tableName+"'["+columnName+"]))").Tables[0];

    for (int r = 0; r < result.Rows.Count; r++)
	{
	    string resultValue = result.Rows[r][0].ToString();
	    if (!resultValue.EndsWith("12:00:00 AM"))
	    {
	        obj.SetAnnotation(annName,"Yes");
	        r=50;
	    }
	    else
	    {
	        obj.SetAnnotation(annName,"No");
	    }
	}
}
