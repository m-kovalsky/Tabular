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
    var tableName = o.Table.Name;
    var hName = o.Name;
    
    Model.Tables[tableName].Hierarchies[hName].RemoveAnnotation("Vertipaq_HierarchyID");
    Model.Tables[tableName].Hierarchies[hName].RemoveAnnotation("Vertipaq_UserHierarchySize");
}

foreach (var o in Model.AllColumns)
{
    var tableName = o.Table.Name;
    var cName = o.Name;
    
    Model.Tables[tableName].Columns[cName].RemoveAnnotation("Vertipaq_ColumnID");
    Model.Tables[tableName].Columns[cName].RemoveAnnotation("Vertipaq_ColumnHierarchySize");
    Model.Tables[tableName].Columns[cName].RemoveAnnotation("Vertipaq_DataSize");
    Model.Tables[tableName].Columns[cName].RemoveAnnotation("Vertipaq_DictionarySize");
    Model.Tables[tableName].Columns[cName].RemoveAnnotation("Vertipaq_Cardinality");
    Model.Tables[tableName].Columns[cName].RemoveAnnotation("Vertipaq_ColumnSize");
}

foreach (var o in Model.Relationships.ToList())
{    
    var rName = o.ID;
    
    Model.Relationships[rName].RemoveAnnotation("Vertipaq_RelationshipID");
    Model.Relationships[rName].RemoveAnnotation("Vertipaq_RelationshipSize");    
}

foreach (var o in Model.Tables.ToList())
{
    var tableName = o.Name;    
    
    Model.Tables[tableName].RemoveAnnotation("Vertipaq_TableID");
    Model.Tables[tableName].RemoveAnnotation("Vertipaq_RowCount");
    Model.Tables[tableName].RemoveAnnotation("Vertipaq_TableSize");
}

foreach (var o in Model.AllPartitions)
{
    var tableName = o.Table.Name;
    var pName = o.Name;
    
    Model.Tables[tableName].Partitions[pName].RemoveAnnotation("Vertipaq_PartitionID");
    Model.Tables[tableName].Partitions[pName].RemoveAnnotation("Vertipaq_PartitionStorageID");
    Model.Tables[tableName].Partitions[pName].RemoveAnnotation("Vertipaq_RecordCount");
    Model.Tables[tableName].Partitions[pName].RemoveAnnotation("Vertipaq_RecordsPerSegment");
    Model.Tables[tableName].Partitions[pName].RemoveAnnotation("Vertipaq_SegmentCount");
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
    var tblName = DMV_Dimensions.Rows[r][0].ToString();
    var recordCount = DMV_Dimensions.Rows[r][1].ToString();
    
    if (tblName != "Measures")
    {
        Model.Tables[tblName].SetAnnotation("Vertipaq_RowCount",recordCount);
    }
}

// Set Relationship IDs
for (int r = 0; r < DMV_Relationships.Rows.Count; r++)
{
    var ID = DMV_Relationships.Rows[r][0].ToString();   
    var relID = DMV_Relationships.Rows[r][1].ToString();    
    
    Model.Relationships[relID].SetAnnotation("Vertipaq_RelationshipID",ID);    
}

// Set Hierarchy IDs
for (int r = 0; r < DMV_Hierarchies.Rows.Count; r++)
{
    var hID = DMV_Hierarchies.Rows[r][0].ToString();
    var tableID = DMV_Hierarchies.Rows[r][1].ToString();
    var hName = DMV_Hierarchies.Rows[r][2].ToString();
    
    foreach (var t in Model.Tables.Where(a => a.GetAnnotation("Vertipaq_TableID") == tableID))
    {
        var tableName = t.Name;
        Model.Tables[tableName].Hierarchies[hName].SetAnnotation("Vertipaq_HierarchyID",hID);
    }        
}

// Set Column IDs
for (int r = 0; r < DMV_Columns.Rows.Count; r++)
{
    var colID = DMV_Columns.Rows[r][0].ToString();
    var tableID = DMV_Columns.Rows[r][1].ToString();
    var colName = DMV_Columns.Rows[r][2].ToString();
    
    foreach (var t in Model.Tables.Where(a => a.GetAnnotation("Vertipaq_TableID") == tableID))
    {
        var tableName = t.Name;
        
        if (colName.StartsWith("RowNumber-") == false)
        {
            Model.Tables[tableName].Columns[colName].SetAnnotation("Vertipaq_ColumnID",colID);
        }
    }
}

// Set Partition IDs
for (int r = 0; r < DMV_Partitions.Rows.Count; r++)
{
    var pID = DMV_Partitions.Rows[r][0].ToString();
    var tableID = DMV_Partitions.Rows[r][1].ToString();
    var pName = DMV_Partitions.Rows[r][2].ToString();
    
    foreach (var t in Model.Tables.Where(a => a.GetAnnotation("Vertipaq_TableID") == tableID))
    {
        var tableName = t.Name;
        
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_PartitionID",pID);        
    }
}


// Set Partition Storage IDs
for (int r = 0; r < DMV_PartitionStorages.Rows.Count; r++)
{
    var psID = DMV_PartitionStorages.Rows[r][0].ToString();
    var pID = DMV_PartitionStorages.Rows[r][1].ToString();    
    
    foreach (var p in Model.AllPartitions.Where(a => a.GetAnnotation("Vertipaq_PartitionID") == pID))
    {
        var tableName = p.Table.Name;
        var pName = p.Name;
        
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_PartitionStorageID",psID);        
    }
}

// Set Partition Stats
for (int r = 0; r < DMV_SegmentMapStorages.Rows.Count; r++)
{
    var psID = DMV_SegmentMapStorages.Rows[r][0].ToString();
    var recordCount = DMV_SegmentMapStorages.Rows[r][1].ToString();    
    var segmentCount = DMV_SegmentMapStorages.Rows[r][2].ToString();    
    var recordsPerSegment = DMV_SegmentMapStorages.Rows[r][3].ToString();    
    
    foreach (var p in Model.AllPartitions.Where(a => a.GetAnnotation("Vertipaq_PartitionStorageID") == psID))
    {
        var tableName = p.Table.Name;
        var pName = p.Name;
        
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_RecordCount",recordCount);
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_SegmentCount",segmentCount);
        Model.Tables[tableName].Partitions[pName].SetAnnotation("Vertipaq_RecordsPerSegment",recordsPerSegment);
    }
}

// Set Dictionary Size
for (int r = 0; r < DMV_StorageTableColumns.Rows.Count; r++)
{
    var tableName = DMV_StorageTableColumns.Rows[r][0].ToString();    
    var colName = DMV_StorageTableColumns.Rows[r][1].ToString();
    var colType = DMV_StorageTableColumns.Rows[r][2].ToString();
    var dictSize = DMV_StorageTableColumns.Rows[r][3].ToString();
      
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
        var hName = o.Name;
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
        var rName = o.ID;
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
        var colName = o.Name;
        int colSize = Convert.ToInt32(Model.Tables[tableName].Columns[colName].GetAnnotation("Vertipaq_ColumnHierarchySize"));
        
        if (usedObj.StartsWith("H$"))
        {
            if (colSize != null)
            {
                colSize = colSize + Convert.ToInt32(usedSize);
            }
            else
            {
                colSize = Convert.ToInt32(usedSize);
            }
        
            Model.Tables[tableName].Columns[colName].SetAnnotation("Vertipaq_ColumnHierarchySize",colSize.ToString());                
        }   
    }  
    
    // Column Data Size
    foreach (var o in Model.Tables[tableName].Columns.Where(a => a.GetAnnotation("Vertipaq_ColumnID") == usedObjID2))
    {
        var colName = o.Name;
        int colSize = Convert.ToInt32(Model.Tables[tableName].Columns[colName].GetAnnotation("Vertipaq_DataSize"));
        
        if (usedObj.StartsWith("H$") == false && usedObj.StartsWith("R$") == false && usedObj.StartsWith("U$") == false)
        {
            if (colSize != null)
            {
                colSize = colSize + Convert.ToInt32(usedSize);
            }
            else
            {
                colSize = Convert.ToInt32(usedSize);
            }
        
            Model.Tables[tableName].Columns[colName].SetAnnotation("Vertipaq_DataSize",colSize.ToString());                
        }
    }   
}

// Set Column & Table Size

int tableSizeCumulative = 0;

foreach (var t in Model.Tables.ToList())
{
    var tableName = t.Name;
    int colSizeCumulative = 0;
    int userHierSizeCumulative = 0;
    int relSizeCumulative = 0;       
    
    foreach (var c in t.Columns.ToList())
    {        
        var colName = c.Name;
        var obj = Model.Tables[tableName].Columns[colName];
        
        int colHierSize = Convert.ToInt32(obj.GetAnnotation("Vertipaq_ColumnHierarchySize"));
        int dataSize = Convert.ToInt32(obj.GetAnnotation("Vertipaq_DataSize"));
        int dictSize = Convert.ToInt32(obj.GetAnnotation("Vertipaq_DictionarySize"));
        
        int colSize = colHierSize + dataSize + dictSize;
        colSizeCumulative = colSizeCumulative + colSize;        
        
        // Set Column Size
        obj.SetAnnotation("Vertipaq_ColumnSize",colSize.ToString());
    }
    
    foreach (var h in t.Hierarchies.ToList())
    {
        var hName = h.Name;
        var obj = Model.Tables[tableName].Hierarchies[hName];
        
        int userHierSize = Convert.ToInt32(obj.GetAnnotation("Vertipaq_UserHierarchySize"));      
        userHierSizeCumulative = userHierSizeCumulative + userHierSize;           
    }
    
    foreach (var r in Model.Relationships.Where(a => a.FromTable.Name == tableName).ToList())
    {
        var rName = r.ID;
        var obj = Model.Relationships[rName];
        
        int relSize = Convert.ToInt32(obj.GetAnnotation("Vertipaq_RelationshipSize"));
        
        relSizeCumulative = relSizeCumulative + relSize;                
    }
    
    int tableSize = colSizeCumulative + userHierSizeCumulative + relSizeCumulative;
    tableSizeCumulative = tableSizeCumulative + tableSize;
    
    // Set Table Size
    Model.Tables[tableName].SetAnnotation("Vertipaq_TableSize",tableSize.ToString());
}

// Set Model Size
Model.SetAnnotation("Vertipaq_ModelSize",tableSizeCumulative.ToString());

// Remove Vertipaq ID Annotations
foreach (var o in Model.AllHierarchies)
{
    var tableName = o.Table.Name;
    var hName = o.Name;
    
    Model.Tables[tableName].Hierarchies[hName].RemoveAnnotation("Vertipaq_HierarchyID");
}

foreach (var o in Model.AllColumns)
{
    var tableName = o.Table.Name;
    var cName = o.Name;
    
    Model.Tables[tableName].Columns[cName].RemoveAnnotation("Vertipaq_ColumnID");
}

foreach (var o in Model.Relationships.ToList())
{    
    var rName = o.ID;
    
    Model.Relationships[rName].RemoveAnnotation("Vertipaq_RelationshipID");    
}

foreach (var o in Model.Tables.ToList())
{
    var tableName = o.Name;    
    
    Model.Tables[tableName].RemoveAnnotation("Vertipaq_TableID");
}

foreach (var o in Model.AllPartitions)
{
    var tableName = o.Table.Name;
    var pName = o.Name;
    
    Model.Tables[tableName].Partitions[pName].RemoveAnnotation("Vertipaq_PartitionID");
    Model.Tables[tableName].Partitions[pName].RemoveAnnotation("Vertipaq_PartitionStorageID");
}
