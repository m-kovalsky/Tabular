// Loop through structured data sources
foreach (var ds in Model.DataSources.Where(a => a.Type.ToString() == "Structured").ToList())
{    
    string strdsName = ds.Name;
    var strDataSource = (Model.DataSources[strdsName] as StructuredDataSource);
    string serverName = strDataSource.Server;
    string dbName = strDataSource.Database;
    string userName = strDataSource.Username;
    
    // Create provider data source
    string provdsName = strdsName + "1";    
    var provDataSource = Model.AddDataSource(provdsName);
    string conn = "Data Source="+serverName+";Initial Catalog="+dbName+";Persist Security Info=True";

    if (userName != null)
    {
        conn = conn + ";User ID="+userName;
    }

    // Update provider data source connection string and provider
    provDataSource.ConnectionString = conn;
    provDataSource.Provider = "System.Data.SqlClient";
    
    // Update partitions to use the new provider data source
    foreach (var t in Model.Tables.ToList())
    {
        foreach (var p in t.Partitions.Where(a => a.DataSource.Name == strdsName).ToList())
        {
            p.DataSource = provDataSource;
        }
    }
    
    // Delete structured data source
    strDataSource.Delete();

    // Rename provider data source
    Model.DataSources[provdsName].Name = strdsName;
}