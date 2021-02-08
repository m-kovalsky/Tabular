// Initial Parameters
string dataSource = @""; // Enter the name of the data source in the model that the Data Dictionary will use.

string serverName = @""; // Enter the server name (where the Data Dictionary table will reside in the data warehouse).
string databaseName = ""; // Enter the database name (where the Data Dictionary table will reside in the data warehouse).
string schemaName = ""; // Enter the schema name (ensure the schema exists).
string dbTableName = ""; // Enter the table name for the Data Dictionary in the data warehouse.

string modelName = Model.Database.Name;

// Error check for parameters
if (!Model.DataSources.Any(a=>a.Name == dataSource))
{
    Error("Must enter a valid 'Data Source' in the dataSource parameter");
    return;
}
if (serverName.Length == 0)
{
    Error("Must enter a valid 'Server' in the serverName parameter");
    return;
}
if (databaseName.Length == 0)
{
    Error("Must enter a valid 'Database' in the databaseName parameter");
    return;
}
if (schemaName.Length == 0)
{
    Error("Must enter a valid 'Schema' in the schemaName parameter");
    return;
}
if (dbTableName.Length == 0)
{
    Error("Must enter a valid 'Table Name' in the dbTableName parameter");
    return;
}
if (modelName.Length == 0)
{
    Error("Must enter a valid 'Model Name' in the modelName parameter");
    return;
}

// Create Data Dictionary table within the model (if it does not already exist)
string ddTableName = "Data Dictionary";
string[] sourceColName = {"ModelName","TableName","ObjectType","ObjectName","HiddenFlag","Description","DisplayFolder","MeasureFormula"};
string[] colName = {"Model","Table","Object Type","Object","Hidden Flag","Description","Display Folder","Measure Formula"};

if (Model.Tables.Any(a => a.Name == ddTableName) == false)
{
    var t = Model.AddTable(ddTableName);    
    t.Partitions[0].DataSource = Model.DataSources[dataSource];
    t.Partitions[0].Query = "SELECT * FROM ["+schemaName+"].["+dbTableName+"]";

    for (int i=0;i<colName.Length; i++)
    {
        var c = t.AddDataColumn(colName[i]);
        c.SourceColumn = sourceColName[i];
        c.DataType = DataType.String;
    }
}

// Create Data Dictionary table within the Data Warehouse
string newLine = Environment.NewLine;
string connectionString = @"Data Source="+serverName+";Initial Catalog="+databaseName+";Integrated Security=True";
string sql = "DROP TABLE IF EXISTS ["+schemaName+"].["+dbTableName+"] "+newLine+
"CREATE TABLE ["+schemaName+"].["+dbTableName+"] " +newLine+
"(" +newLine+
" [ModelName] VARCHAR(100)" +newLine+
",[TableName] VARCHAR(200)" +newLine+
",[ObjectType] VARCHAR(30)" +newLine+
",[ObjectName] VARCHAR(250)" +newLine+
",[HiddenFlag] VARCHAR(10)" +newLine+
",[Description] VARCHAR(MAX)" +newLine+
",[DisplayFolder] VARCHAR(150)" +newLine+
",[MeasureFormula] VARCHAR(MAX)" +newLine+
")";

System.Data.SqlClient.SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(sql, con);

cmd.Connection.Open();
cmd.ExecuteNonQuery();

// Set up the insert statement for sending the metadata to the data warehouse
string insertSQL;
var sb_InsertSQL = new System.Text.StringBuilder();

sb_InsertSQL.Append("INSERT INTO ["+schemaName+"].["+dbTableName+"]");
sb_InsertSQL.Append(Environment.NewLine);

// Extract model metadata in the data dictionary format
foreach (var t in Model.Tables.Where(a => a.ObjectType.ToString() != "CalculationGroupTable" && a.Name != ddTableName).OrderBy(a => a.Name).ToList())
{
    string tableName = t.Name;
    string tableDesc = t.Description.Replace("'","''");
    string objectType = "Table";
    string hiddenFlag;                 
    string expr;

    if (t.IsHidden)
    {
        hiddenFlag = "Yes";
    }
    else
    {
        hiddenFlag = "No";
    }
    
    if (t.SourceType.ToString() == "Calculated")
    {
        expr = (Model.Tables[tableName] as CalculatedTable).Expression;
 
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ");
        expr = expr.Replace("\t"," ");
        expr = expr.Replace("'","''");
        
        sb_InsertSQL.Append("SELECT"+"'"+modelName+"','"+tableName+"','"+objectType+"','"+tableName+"','"+hiddenFlag+"','"+tableDesc+"','"+"','"+expr+"' UNION ALL ");                        
    }
    else
    {
        sb_InsertSQL.Append("SELECT"+"'"+modelName+"','"+tableName+"','"+objectType+"','"+tableName+"','"+hiddenFlag+"','"+tableDesc+"','"+"','"+"***N/A***' UNION ALL ");             
    }
    
    sb_InsertSQL.Append(Environment.NewLine);
    
    foreach (var o in t.Columns.OrderBy(a => a.Name).ToList())
    {
        string objectName = o.Name;
        string objectDesc = o.Description.Replace("'","''");
        string objectDF = o.DisplayFolder.Replace("'","''");
        objectType = "Attribute";
        
        if (o.IsHidden)
        {
            hiddenFlag = "Yes";
        }
        else
        {
            hiddenFlag = "No";
        }
        
        if (o.Type.ToString() == "Calculated")
        {
            expr = (Model.Tables[tableName].Columns[objectName] as CalculatedColumn).Expression;
            
            // Remove tabs and new lines
            expr = expr.Replace("\n"," ");
            expr = expr.Replace("\t"," ");
            expr = expr.Replace("'","''");
            sb_InsertSQL.Append("SELECT"+"'"+modelName+"','"+tableName+"','"+objectType+"','"+objectName+"','"+hiddenFlag+"','"+objectDesc+"','"+objectDF+"','"+expr+"' UNION ALL ");                
            
        }
        else
        {
            sb_InsertSQL.Append("SELECT"+"'"+modelName+"','"+tableName+"','"+objectType+"','"+objectName+"','"+hiddenFlag+"','"+objectDesc+"','"+objectDF+"','***N/A***' UNION ALL ");        
        }
        
        sb_InsertSQL.Append(Environment.NewLine);        
    }
    
    foreach (var o in t.Measures.OrderBy(a => a.Name).ToList())
    {
        string objectName = o.Name;
        string objectDesc = o.Description.Replace("'","''");
        string objectDF = o.DisplayFolder.Replace("'","''");
        objectType = "Measure";
        expr = o.Expression;                    
        
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ");
        expr = expr.Replace("\t"," ");
        expr = expr.Replace("'","''");
        
        if (o.IsHidden)
        {
            hiddenFlag = "Yes";
        }
        else
        {
            hiddenFlag = "No";
        }
        
        sb_InsertSQL.Append("SELECT"+"'"+modelName+"','"+tableName+"','"+objectType+"','"+objectName+"','"+hiddenFlag+"','"+objectDesc+"','"+objectDF+"','"+expr+"' UNION ALL ");        
        sb_InsertSQL.Append(Environment.NewLine);
    }
    
    foreach (var o in t.Hierarchies.OrderBy(a => a.Name).ToList())
    {
        string objectName = o.Name;
        string objectDesc = o.Description.Replace("'","''");
        string objectDF = o.DisplayFolder.Replace("'","''");
        objectType = "Hierarchy";
        
        if (o.IsHidden)
        {
            hiddenFlag = "Yes";
        }
        else
        {
            hiddenFlag = "No";
        }
        
        sb_InsertSQL.Append("SELECT"+"'"+modelName+"','"+tableName+"','"+objectType+"','"+objectName+"','"+hiddenFlag+"','"+objectDesc+"','"+objectDF+"','***N/A***' UNION ALL ");
        sb_InsertSQL.Append(Environment.NewLine);
    }            
}

foreach (var o in Model.CalculationGroups.ToList())
{
    string tableName = o.Name;
    string tableDesc = o.Description.Replace("'","''");
    string hiddenFlag;
    
    if (o.IsHidden)
    {
        hiddenFlag = "Yes";
    }
    else
    {
        hiddenFlag = "No";
    }
    
    sb_InsertSQL.Append("SELECT"+"'"+modelName+"','"+tableName+"','Calculation Group','"+tableName+"','"+hiddenFlag+"','"+tableDesc+"','','***N/A***' UNION ALL ");    
    sb_InsertSQL.Append(Environment.NewLine);
    
    foreach (var i in o.CalculationItems.ToList())
    {        
        string objectName = i.Name;
        string objectDesc = i.Description.Replace("'","''");
        string expr = i.Expression;            
        
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ");
        expr = expr.Replace("\t"," ");
        expr = expr.Replace("'","''");
        
        sb_InsertSQL.Append("SELECT"+"'"+modelName+"','"+tableName+"','Calculation Item','"+objectName+"','No','"+objectDesc+"','','"+expr+"' UNION ALL ");                
        sb_InsertSQL.Append(Environment.NewLine);
    }
} 

// Remove the extra comma 
insertSQL = sb_InsertSQL.ToString().Trim();
insertSQL = insertSQL.Substring(0,insertSQL.Length-10);

// Insert the data dictionary metadata into the data warehouse
System.Data.SqlClient.SqlCommand cmdInsert = new System.Data.SqlClient.SqlCommand(insertSQL, con);

cmdInsert.ExecuteNonQuery();
cmdInsert.Connection.Close();
cmd.Connection.Close();
con.Close();