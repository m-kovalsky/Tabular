// Select the columns from your detail fact table that you would like in the agg table. 
// Only specify foreign keys or columns used for measures.

var tableName = Selected.Table.Name;
string aggSuffix = "_Agg";
string aggTableName = tableName + aggSuffix;

// Ensure no duplication of agg table
if (!Model.Tables.Any(y => y.Name == aggTableName))
{
    // Create agg table
    var aggTable = Model.AddTable(aggTableName);
    aggTable.IsHidden = true;

    int pCount = Model.Tables[tableName].Partitions.Count();
    var dataSource = Model.Tables[tableName].Source;

    // For multi-partitioned tables
    if (pCount > 1)
    {
        // Add partitions
        foreach(var p in Model.Tables[tableName].Partitions.ToList())
        {
            var aggPartitionName = p.Name.Replace(tableName,aggTableName);
            var aggQuery = p.Query.Replace(tableName,aggTableName);
            
            if (p.SourceType.ToString() == "M")
            {
                var pName = Model.Tables[aggTableName].AddMPartition(aggPartitionName);
                pName.Query = aggQuery;
            }
            else
            {
                var pName = Model.Tables[aggTableName].AddPartition(aggPartitionName);
                pName.DataSource = Model.DataSources[dataSource];
                pName.Query = aggQuery;
            }
        }
        
        // Remove default partition
        Model.Tables[aggTableName].Partitions[aggTableName].Delete();
    }

    // For single-partition tables
    else
    {
        var par = Model.Tables[tableName].Partitions[tableName];
        var aggPar = Model.Tables[aggTableName].Partitions[aggTableName];
        
        if (aggPar.SourceType.ToString() == "M")
        {
            aggPar.Query = par.Query;
        }        
        else
        {
            // Update Data Source
            aggPar.DataSource = Model.DataSources[dataSource];
            
            // Update Query
            aggPar.Query = par.Query.Replace(tableName,aggTableName);
        }
    }

    foreach (var c in Selected.Columns)
    {
        // Add Columns
        string columnName = c.Name;
        bool hide = c.IsHidden;
        var colfs = c.FormatString;
        var dt = c.DataType;
        var sourceColumn = (Model.Tables[tableName].Columns[columnName] as DataColumn).SourceColumn;
        
        // Add Column Properties
        var obj = Model.Tables[aggTableName].AddDataColumn(columnName);
        obj.SourceColumn = sourceColumn;
        obj.IsHidden = hide;
        obj.FormatString = colfs;
        obj.DataType = dt;
        
        // Add Relationships
        foreach(var r in Model.Relationships.ToList().Where(a=> a.FromTable == Model.Tables[tableName] && a.FromColumn == Model.Tables[tableName].Columns[columnName]))
        {
            var addRel = Model.AddRelationship();
            addRel.FromColumn = Model.Tables[aggTableName].Columns[columnName];
            addRel.ToColumn = Model.Tables[r.ToTable.Name].Columns[r.ToColumn.Name];
            addRel.CrossFilteringBehavior = r.CrossFilteringBehavior;
            addRel.SecurityFilteringBehavior = r.SecurityFilteringBehavior;
            addRel.IsActive = r.IsActive;
            Model.Tables[aggTableName].Columns[columnName].SetAnnotation(aggTableName,"ForeignKey");
            Model.Tables[tableName].Columns[columnName].SetAnnotation(aggTableName,"ForeignKey");
        }
        
        // For non-key columns, create agg measures
        if ( Model.Tables[aggTableName].Columns[columnName].GetAnnotation(aggTableName) == null)
        {
            foreach (var x in Model.Tables[tableName].Columns[columnName].ReferencedBy.OfType<Measure>().ToList())
            {
                var newMeasureName = x.Name + aggSuffix;
                var measureDAX = x.Expression;
                var newDAX = measureDAX.Replace(tableName + "[" + columnName + "]",aggTableName + "[" + columnName + "]");
                newDAX = newDAX.Replace("'" + tableName + "'" + "[" + columnName + "]","'" + aggTableName + "'" + "[" + columnName + "]");
                var fs = x.FormatString;
                var df = x.DisplayFolder;
                var k = x.KPI;
                
                // Create agg measure, format same as non-agg measure
                var newMeasure = Model.Tables[aggTableName].AddMeasure(newMeasureName);
                newMeasure.Expression = FormatDax(newDAX);
                newMeasure.IsHidden = true;
                newMeasure.FormatString = fs;
                newMeasure.DisplayFolder = df;
                newMeasure.KPI = k;
                
                // Add new measures to respective perspectives
                foreach (var p in Model.Perspectives.ToList())
                {
                    foreach (var mea in Model.AllMeasures.Where(a=> a.Name == x.Name))
                    {
                        bool inPer = mea.InPerspective[p];
                        newMeasure.InPerspective[p] = inPer;
                        
                        // Set Annotations for base measures
                        mea.SetAnnotation(aggTableName,"BaseMeasure");
                    }
                }
                
                // Set annotation denoting column as an agg column
                Model.Tables[aggTableName].Columns[columnName].SetAnnotation(aggTableName,"AggColumn");
                Model.Tables[tableName].Columns[columnName].SetAnnotation(aggTableName,"AggColumn");
            }
        }
        
        // Add columns to respective perspective(s)
        foreach (var p in Model.Perspectives.ToList())
        {
            bool inPersp = c.InPerspective[p];            
            obj.InPerspective[p] = inPersp;
        }
    }

    // Initialize DAX Statement string for Agg-check measure
    var sb = new System.Text.StringBuilder();
    sb.Append("IF (");
    
    // Create ISCROSSFILTERED Statement
    foreach (var c in Model.Tables[tableName].Columns.Where(b => b.GetAnnotation(aggTableName) != "ForeignKey" && b.GetAnnotation(aggTableName) != "AggColumn").ToList())
    {
        foreach(var r in Model.Relationships.ToList().Where(a=> a.FromTable == Model.Tables[tableName] && a.FromColumn == Model.Tables[tableName].Columns[c.Name]))
        {
            sb.Append("ISCROSSFILTERED('"+r.ToTable.Name+"') || ");
            Model.Tables[tableName].Columns[c.Name].SetAnnotation(aggTableName,"ForeignKeyNotInAgg");
        }
    }   
    
    // Create ISFILTERED Statement    
    foreach (var c in Model.Tables[tableName].Columns.Where(b => b.GetAnnotation(aggTableName) == null).ToList())
    {
        sb.Append("ISFILTERED('"+tableName+"'["+c.Name+"]) || ");
    }
    
    string dax = sb.ToString(0,sb.Length - 3) + ",0,1)";

    var m = Model.Tables[aggTableName].AddMeasure(aggTableName+"check");
    m.Expression = FormatDax(dax);
    m.IsHidden = true;
    
    // Add Agg-check measure to respective perspective(s)
    foreach (var t in Model.Tables.Where (a => a.Name == tableName))
    {        
        foreach (var p in Model.Perspectives.ToList())
        {
            bool inPersp = t.InPerspective[p];            
            m.InPerspective[p] = inPersp;
        }
    }
}

// Update non-agg measures to switch between agg & non-agg
foreach (var a in Model.AllMeasures.Where(a => a.GetAnnotation(aggTableName) == "BaseMeasure").ToList())
{
    a.Expression = FormatDax("IF([" + aggTableName + "check] = 1,[" + a.Name + aggSuffix +"],"+a.Expression+")");
}