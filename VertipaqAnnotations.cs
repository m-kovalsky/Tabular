// Store DMV Queries as Data Tables
var DMV_Tables = ExecuteDax("SELECT [ID],[Name] FROM $SYSTEM.TMSCHEMA_TABLES").Tables[0];
var DMV_Dimensions = ExecuteDax("SELECT [DIMENSION_NAME], [DIMENSION_CARDINALITY] FROM $SYSTEM.MDSCHEMA_DIMENSIONS").Tables[0];
var DMV_Relationships = ExecuteDax("SELECT [ID],[Name] FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS").Tables[0];
var DMV_Hierarchies = ExecuteDax("SELECT [ID], [TableID], [Name] FROM $SYSTEM.TMSCHEMA_HIERARCHIES").Tables[0];
var DMV_Columns = ExecuteDax("SELECT [ID],[TableID],[ExplicitName] FROM $SYSTEM.TMSCHEMA_COLUMNS").Tables[0];
var DMV_Partitions = ExecuteDax("SELECT [ID],[TableID],[Name] FROM $SYSTEM.TMSCHEMA_PARTITIONS").Tables[0];
var DMV_PartitionStorages = ExecuteDax("SELECT [ID],[PartitionID] FROM $SYSTEM.TMSCHEMA_PARTITION_STORAGES").Tables[0];
var DMV_SegmentMapStorages = ExecuteDax("SELECT [PartitionStorageID],[RecordCount],[SegmentCount],[RecordsPerSegment] FROM $SYSTEM.TMSCHEMA_SEGMENT_MAP_STORAGES").Tables[0];
var DMV_StorageTableColumns = ExecuteDax("SELECT [DIMENSION_NAME],[ATTRIBUTE_NAME],[COLUMN_TYPE],[DICTIONARY_SIZE] FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS").Tables[0];
var DMV_StorageTables = ExecuteDax("SELECT [DIMENSION_NAME],[TABLE_ID],[ROWS_COUNT] FROM $SYSTEM.DISCOVER_STORAGE_TABLES").Tables[0];
var DMV_ColumnSegments = ExecuteDax("SELECT [DIMENSION_NAME],[TABLE_ID],[COLUMN_ID],[USED_SIZE] FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS").Tables[0];

// Remove Existing Vertipaq Annotations
Model.RemoveAnnotation("Vertipaq_ModelSize");

foreach (var o in Model.AllHierarchies)
{   
    o.RemoveAnnotation("Vertipaq_HierarchyID");
    o.RemoveAnnotation("Vertipaq_UserHierarchySize");
}

foreach (var o in Model.AllColumns)
{   
    o.RemoveAnnotation("Vertipaq_ColumnID");
    o.RemoveAnnotation("Vertipaq_ColumnHierarchySize");
    o.RemoveAnnotation("Vertipaq_DataSize");
    o.RemoveAnnotation("Vertipaq_DictionarySize");
    o.RemoveAnnotation("Vertipaq_Cardinality");
    o.RemoveAnnotation("Vertipaq_ColumnSize");
    o.RemoveAnnotation("Vertipaq_ColumnSizePctOfTable");
    o.RemoveAnnotation("Vertipaq_ColumnSizePctOfModel");
}

foreach (var o in Model.Relationships.ToList())
{    
    o.RemoveAnnotation("Vertipaq_RelationshipID");
    o.RemoveAnnotation("Vertipaq_RelationshipSize");
    o.RemoveAnnotation("Vertipaq_MaxFromCardinality");
    o.RemoveAnnotation("Vertipaq_MaxToCardinality");
    o.RemoveAnnotation("Vertipaq_RIViolationInvalidRows");
}

foreach (var o in Model.Tables.ToList())
{
    o.RemoveAnnotation("Vertipaq_TableID");
    o.RemoveAnnotation("Vertipaq_RowCount");
    o.RemoveAnnotation("Vertipaq_TableSize");
    o.RemoveAnnotation("Vertipaq_TableSizePctOfModel");
}

foreach (var o in Model.AllPartitions)
{
    o.RemoveAnnotation("Vertipaq_PartitionID");
    o.RemoveAnnotation("Vertipaq_PartitionStorageID");
    o.RemoveAnnotation("Vertipaq_RecordCount");
    o.RemoveAnnotation("Vertipaq_RecordsPerSegment");
    o.RemoveAnnotation("Vertipaq_SegmentCount");
}

// Set Table IDs
for (int r = 0; r < DMV_Tables.Rows.Count; r++)
{
    string tblID = DMV_Tables.Rows[r][0].ToString();
    string tblName = DMV_Tables.Rows[r][1].ToString();
    
    Model.Tables[tblName].SetAnnotation("Vertipaq_TableID",tblID);         
}

// Set Table Row Counts
for (int r = 0; r < DMV_Dimensions.Rows.Count; r++)
{
    string tblName = DMV_Dimensions.Rows[r][0].ToString();
    string recordCount = DMV_Dimensions.Rows[r][1].ToString();
    
    if (tblName != "Measures")
    {
        Model.Tables[tblName].SetAnnotation("Vertipaq_RowCount",recordCount);
    }
}

// Set Relationship IDs
for (int r = 0; r < DMV_Relationships.Rows.Count; r++)
{
    string ID = DMV_Relationships.Rows[r][0].ToString();   
    string relID = DMV_Relationships.Rows[r][1].ToString();    
    
    Model.Relationships[relID].SetAnnotation("Vertipaq_RelationshipID",ID);    
}

// Set Hierarchy IDs
for (int r = 0; r < DMV_Hierarchies.Rows.Count; r++)
{
    string hID = DMV_Hierarchies.Rows[r][0].ToString();
    string tableID = DMV_Hierarchies.Rows[r][1].ToString();
    string hName = DMV_Hierarchies.Rows[r][2].ToString();
    
    foreach (var t in Model.Tables.Where(a => a.GetAnnotation("Vertipaq_TableID") == tableID))
    {
        string tableName = t.Name;
        Model.Tables[tableName].Hierarchies[hName].SetAnnotation("Vertipaq_HierarchyID",hID);
    }        
}

// Set Column IDs
for (int r = 0; r < DMV_Columns.Rows.Count; r++)
{
    string colID = DMV_Columns.Rows[r][0].ToString();
    string tableID = DMV_Columns.Rows[r][1].ToString();
    string colName = DMV_Columns.Rows[r][2].ToString();
    
    foreach (var t in Model.Tables.Where(a => a.GetAnnotation("Vertipaq_TableID") == tableID))
    {
        string tableName = t.Name;
        
        if (colName.StartsWith("RowNumber-") == false && colName != "")
        {
            Model.Tables[tableName].Columns[colName].SetAnnotation("Vertipaq_ColumnID",colID);
        }
    }
}

// Set Partition IDs
for (int r = 0; r < DMV_Partitions.Rows.Count; r++)
{
    string pID = DMV_Partitions.Rows[r][0].ToString();
    string tableID = DMV_Partitions.Rows[r][1].ToString();
    string pName = DMV_Partitions.Rows[r][2].ToString();
    
    foreach (var t in Model.Tables.Where(a => a.GetAnnotation("Vertipaq_TableID") == tableID))
    {
        string tableName = t.Name;
        
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_PartitionID",pID);        
    }
}


// Set Partition Storage IDs
for (int r = 0; r < DMV_PartitionStorages.Rows.Count; r++)
{
    string psID = DMV_PartitionStorages.Rows[r][0].ToString();
    string pID = DMV_PartitionStorages.Rows[r][1].ToString();    
    
    foreach (var p in Model.AllPartitions.Where(a => a.GetAnnotation("Vertipaq_PartitionID") == pID))
    {
        string tableName = p.Table.Name;
        string pName = p.Name;
        
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_PartitionStorageID",psID);        
    }
}

// Set Partition Stats
for (int r = 0; r < DMV_SegmentMapStorages.Rows.Count; r++)
{
    string psID = DMV_SegmentMapStorages.Rows[r][0].ToString();
    string recordCount = DMV_SegmentMapStorages.Rows[r][1].ToString();    
    string segmentCount = DMV_SegmentMapStorages.Rows[r][2].ToString();    
    string recordsPerSegment = DMV_SegmentMapStorages.Rows[r][3].ToString();    
    
    foreach (var p in Model.AllPartitions.Where(a => a.GetAnnotation("Vertipaq_PartitionStorageID") == psID))
    {
        string tableName = p.Table.Name;
        string pName = p.Name;
        
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_RecordCount",recordCount);
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_SegmentCount",segmentCount);
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_RecordsPerSegment",recordsPerSegment);
    }
}

// Set Dictionary Size
for (int r = 0; r < DMV_StorageTableColumns.Rows.Count; r++)
{
    string tableName = DMV_StorageTableColumns.Rows[r][0].ToString();    
    string colName = DMV_StorageTableColumns.Rows[r][1].ToString();
    string colType = DMV_StorageTableColumns.Rows[r][2].ToString();
    string dictSize = DMV_StorageTableColumns.Rows[r][3].ToString();
      
    if (colType == "BASIC_DATA" && colName.StartsWith("RowNumber-") == false)
    {
        Model.Tables[tableName].Columns[colName].SetAnnotation("Vertipaq_DictionarySize",dictSize);
    }
}

// Set Column Row Counts
for (int r = 0; r < DMV_StorageTables.Rows.Count; r++)
{
    string tableName = DMV_StorageTables.Rows[r][0].ToString();    
    string usedColumn = DMV_StorageTables.Rows[r][1].ToString();    
    string rowCount = DMV_StorageTables.Rows[r][2].ToString();    
    int lastInd = usedColumn.LastIndexOf("(");
    string usedColumnID = usedColumn.Substring(lastInd+1,usedColumn.Length - lastInd - 2);
    
    
    foreach (var c in Model.Tables[tableName].Columns.Where(a => a.GetAnnotation("Vertipaq_ColumnID") == usedColumnID))
    {
        var colName = c.Name;
        Model.Tables[tableName].Columns[colName].SetAnnotation("Vertipaq_Cardinality",rowCount);
    }    
}

// User Hierarchy Size
for (int r = 0; r < DMV_ColumnSegments.Rows.Count; r++)
{
    string tableName = DMV_ColumnSegments.Rows[r][0].ToString();    
    string usedObj = DMV_ColumnSegments.Rows[r][1].ToString();    
    string usedCol = DMV_ColumnSegments.Rows[r][2].ToString();    
    string usedSize = DMV_ColumnSegments.Rows[r][3].ToString();    
    
    int lastInd = usedObj.LastIndexOf("(");
    string usedObjID = usedObj.Substring(lastInd+1,usedObj.Length - lastInd - 2);    
    
    int lastInd2 = usedCol.LastIndexOf("(");
    string usedObjID2 = usedCol.Substring(lastInd2+1,usedCol.Length - lastInd2 - 2);    
    
    // User Hierarchy Size
    foreach (var o in Model.Tables[tableName].Hierarchies.Where(a => a.GetAnnotation("Vertipaq_HierarchyID") == usedObjID))
    {
        string hName = o.Name;
        int hSize = Convert.ToInt32(Model.Tables[tableName].Hierarchies[hName].GetAnnotation("Vertipaq_UserHierarchySize"));
        
        if (usedObj.StartsWith("U$"))
        {
            if (hSize != null)
            {
                hSize = hSize + Convert.ToInt32(usedSize);
            }
            else
            {
                hSize = Convert.ToInt32(usedSize);
            }
        
            Model.Tables[tableName].Hierarchies[hName].SetAnnotation("Vertipaq_UserHierarchySize",hSize.ToString());                
        }   
    }    
    
    // Relationship Size
    foreach (var o in Model.Relationships.Where(a => a.GetAnnotation("Vertipaq_RelationshipID") == usedObjID))
    {
        string rName = o.ID;
        int rSize = Convert.ToInt32(Model.Relationships[rName].GetAnnotation("Vertipaq_RelationshipSize"));
        
        if (usedObj.StartsWith("R$"))
        {
            if (rSize != null)
            {
                rSize = rSize + Convert.ToInt32(usedSize);
            }
            else
            {
                rSize = Convert.ToInt32(usedSize);
            }
        
            Model.Relationships[rName].SetAnnotation("Vertipaq_RelationshipSize",rSize.ToString());                
        }
    }
    
    // Column Hierarchy Size
    foreach (var o in Model.Tables[tableName].Columns.Where(a => a.GetAnnotation("Vertipaq_ColumnID") == usedObjID))
    {
        string colName = o.Name;
        long colSize = Convert.ToInt64(Model.Tables[tableName].Columns[colName].GetAnnotation("Vertipaq_ColumnHierarchySize"));
        
        if (usedObj.StartsWith("H$"))
        {
            if (colSize != null)
            {
                colSize = colSize + Convert.ToInt32(usedSize);
            }
            else
            {
                colSize = Convert.ToInt64(usedSize);
            }
        
            Model.Tables[tableName].Columns[colName].SetAnnotation("Vertipaq_ColumnHierarchySize",colSize.ToString());                
        }   
    }  
    
    // Column Data Size
    foreach (var o in Model.Tables[tableName].Columns.Where(a => a.GetAnnotation("Vertipaq_ColumnID") == usedObjID2))
    {
        string colName = o.Name;
        long colSize = Convert.ToInt64(Model.Tables[tableName].Columns[colName].GetAnnotation("Vertipaq_DataSize"));
        
        if (usedObj.StartsWith("H$") == false && usedObj.StartsWith("R$") == false && usedObj.StartsWith("U$") == false)
        {
            if (colSize != null)
            {
                colSize = colSize + Convert.ToInt64(usedSize);
            }
            else
            {
                colSize = Convert.ToInt64(usedSize);
            }
        
            Model.Tables[tableName].Columns[colName].SetAnnotation("Vertipaq_DataSize",colSize.ToString());                
        }
    }   
}

// Set Column & Table Size
long tableSizeCumulative = 0;

foreach (var t in Model.Tables.ToList())
{
    string tableName = t.Name;
    long colSizeCumulative = 0;
    long userHierSizeCumulative = 0;
    long relSizeCumulative = 0;       
    
    foreach (var c in t.Columns.ToList())
    {        
        string colName = c.Name;
        var obj = Model.Tables[tableName].Columns[colName];
        
        long colHierSize = Convert.ToInt64(obj.GetAnnotation("Vertipaq_ColumnHierarchySize"));
        long dataSize = Convert.ToInt64(obj.GetAnnotation("Vertipaq_DataSize"));
        long dictSize = Convert.ToInt64(obj.GetAnnotation("Vertipaq_DictionarySize"));
        
        long colSize = colHierSize + dataSize + dictSize;
        colSizeCumulative = colSizeCumulative + colSize;        
        
        // Set Column Size
        obj.SetAnnotation("Vertipaq_ColumnSize",colSize.ToString());
    }
    
    foreach (var h in t.Hierarchies.ToList())
    {
        string hName = h.Name;
        var obj = Model.Tables[tableName].Hierarchies[hName];
        
        long userHierSize = Convert.ToInt32(obj.GetAnnotation("Vertipaq_UserHierarchySize"));      
        userHierSizeCumulative = userHierSizeCumulative + userHierSize;           
    }
    
    foreach (var r in Model.Relationships.Where(a => a.FromTable.Name == tableName).ToList())
    {
        string rName = r.ID;
        var obj = Model.Relationships[rName];
        
        long relSize = Convert.ToInt32(obj.GetAnnotation("Vertipaq_RelationshipSize"));
        
        relSizeCumulative = relSizeCumulative + relSize;                
    }
    
    long tableSize = colSizeCumulative + userHierSizeCumulative + relSizeCumulative;
    tableSizeCumulative = tableSizeCumulative + tableSize;
    
    // Set Table Size
    Model.Tables[tableName].SetAnnotation("Vertipaq_TableSize",tableSize.ToString());
}

// Set Model Size
Model.SetAnnotation("Vertipaq_ModelSize",tableSizeCumulative.ToString());

// Set Max From/To Cardinality, Referential Integrity Violations
foreach (var r in Model.Relationships.ToList())
{
    string rName = r.ID;
    string fromTbl = r.FromTable.Name;
    string fromCol = r.FromColumn.Name;
    string toTbl = r.ToTable.Name;
    string toCol = r.ToColumn.Name;
    //var obj = Model.Relationships[rName];
    bool act = r.IsActive;
    string fromTableFull = r.FromTable.DaxObjectFullName;    
    string fromObject = r.FromColumn.DaxObjectFullName;
    string toObject = r.ToColumn.DaxObjectFullName;
    string dax;
    
    // Set Max From/To Cardinality
    string fromCard = Model.Tables[fromTbl].Columns[fromCol].GetAnnotation("Vertipaq_Cardinality");
    string toCard = Model.Tables[toTbl].Columns[toCol].GetAnnotation("Vertipaq_Cardinality");
    
    r.SetAnnotation("Vertipaq_MaxFromCardinality",fromCard);
    r.SetAnnotation("Vertipaq_MaxToCardinality",toCard);   

    // Set Referential Integrity Violations
    if (act)
    {
        dax = "SUMMARIZECOLUMNS(\"test\",CALCULATE(COUNTROWS("+fromTableFull+"),ISBLANK("+toObject+")))";
    }
    else
    {
        dax = "SUMMARIZECOLUMNS(\"test\",CALCULATE(COUNTROWS("+fromTableFull+"),USERELATIONSHIP("+fromObject+","+toObject+"),ISBLANK("+toObject+")))";
    }
    
    var daxResult = EvaluateDax(dax);
    string blankRowCount = daxResult.ToString();
    
    if (blankRowCount != "Table")
    {
        r.SetAnnotation("Vertipaq_RIViolationInvalidRows",blankRowCount);        
    }
    else
    {
        r.SetAnnotation("Vertipaq_RIViolationInvalidRows","0");
    }
}

// Percent of Table and Model
float modelSize = Convert.ToInt64(Model.GetAnnotation("Vertipaq_ModelSize"));

foreach (var t in Model.Tables.ToList())
{
    string tableName = t.Name;
    var obj = Model.Tables[tableName];
    
    float tableSize = Convert.ToInt64(obj.GetAnnotation("Vertipaq_TableSize"));
    double tblpct = Math.Round(tableSize / modelSize,3);
        
    obj.SetAnnotation("Vertipaq_TableSizePctOfModel",tblpct.ToString());
    
    foreach (var c in t.Columns.ToList())
    {
        string colName = c.Name;
        var col = Model.Tables[tableName].Columns[colName];
        
        float colSize = Convert.ToInt64(col.GetAnnotation("Vertipaq_ColumnSize"));
        double colpctTbl = Math.Round(colSize / tableSize,3);
        double colpctModel = Math.Round(colSize / modelSize,3);
        
        col.SetAnnotation("Vertipaq_ColumnSizePctOfTable",colpctTbl.ToString());
        col.SetAnnotation("Vertipaq_ColumnSizePctOfModel",colpctModel.ToString());
    }
}

// Remove Vertipaq ID Annotations
foreach (var o in Model.AllHierarchies)
{   
    o.RemoveAnnotation("Vertipaq_HierarchyID");
}

foreach (var o in Model.AllColumns)
{
    o.RemoveAnnotation("Vertipaq_ColumnID");
}

foreach (var o in Model.Relationships.ToList())
{    
    o.RemoveAnnotation("Vertipaq_RelationshipID");     
}

foreach (var o in Model.Tables.ToList())
{
    o.RemoveAnnotation("Vertipaq_TableID");
}

foreach (var o in Model.AllPartitions)
{
    o.RemoveAnnotation("Vertipaq_PartitionID");
    o.RemoveAnnotation("Vertipaq_PartitionStorageID");
}