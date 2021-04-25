// Split Date/Time recommendation
string annName = "DateTimeWithHourMinSec";
foreach (var c in Model.AllColumns.Where(a => a.DataType == DataType.DateTime))
{
    string objName = c.DaxObjectFullName;
    var result = ExecuteDax("EVALUATE TOPN(5,SUMMARIZECOLUMNS("+objName+"))").Tables[0];

    for (int r = 0; r < result.Rows.Count; r++)
    {
        string resultValue = result.Rows[r][0].ToString();
        if (!resultValue.EndsWith("12:00:00 AM"))
        {
            c.SetAnnotation(annName,"Yes");
            r=50;
        }
        else
        {
            c.SetAnnotation(annName,"No");
        }
    }
}