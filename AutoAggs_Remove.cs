var tableName = Selected.Table.Name; // or specify table name directly i.e. "Revenue"
var aggTableName = string.Empty;
string aggSuffix = "_Agg";

// Ensure code works if detail or agg table is specified in the tableName parameter
if (tableName.EndsWith(aggSuffix))
{
    aggTableName = tableName;
    tableName = tableName.Substring(0,tableName.Length - 4);
}
else
{
    aggTableName = tableName + aggSuffix;
}

// Run if the agg table exists
if (Model.Tables.Where(a => a.Name == aggTableName).Count() != 0)
{
    // Delete agg table
    Model.Tables[aggTableName].Delete();
    
    foreach (var m in Model.AllMeasures.ToList())//.Where(a => a.Name == tableName).ToList()) // perhaps remove where clause
    {
        if (m.GetAnnotation(aggTableName) == "BaseMeasure")
        {
            var expr = m.Expression;
            var aggMeasureName = m.Name+aggSuffix;
            int aggNameLen = aggMeasureName.Length;
            int startPoint = expr.IndexOf(aggMeasureName)+aggNameLen+2;
            var newExpr = expr.Substring(startPoint,expr.Length - startPoint - 1);
            
            // Update DAX for base measures
            m.Expression = FormatDax(newExpr);

            // Remove measure annotations
            m.RemoveAnnotation(aggTableName);
         
        }
    }
    
    // Remove column annotations
    foreach (var c in Model.Tables[tableName].Columns.ToList())
    {
        c.RemoveAnnotation(aggTableName);
    }   
}


