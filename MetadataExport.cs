var folderName = @"C:\Users\mikova\Desktop\Metadata"; // Update this location to the destination folder on your computer

/******************************DATA SOURCES****************************/

var sb = new System.Text.StringBuilder();

// Headers
sb.Append("DataSource" + '\t' + "ConnectionString" + '\t' + "Provider" + '\t' + "MaxConnections");
sb.Append(Environment.NewLine);

foreach (var o in Model.DataSources.Where(a => a.Type.ToString() == "Provider").ToList())
{
    var ds = (Model.DataSources[o.Name] as ProviderDataSource);
    string n = o.Name;
    string conn = ds.ConnectionString;
    string mc = ds.MaxConnections.ToString();
    string prov = ds.Provider;
    
     sb.Append(n + '\t' + conn + '\t' + prov + '\t' + mc);
}

System.IO.File.WriteAllText(folderName + @"\DataSources.txt", sb.ToString());

/**********************************************************************/
/*********************************TABLES*******************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("TableName" + '\t' + "PartitionName" + '\t' + "DataSource" + '\t' + "Mode" + '\t' +
          "DataCategory" + '\t' + "Description");
sb.Append(Environment.NewLine);

foreach (var o in Model.Tables.ToList())
{
    foreach (var p in o.Partitions.ToList())
    {

        var q = Model.Tables[o.Name].Partitions[p.Name].Query;
        var m = Model.Tables[o.Name].Partitions[p.Name].Mode;
        var dc = string.Empty;
    
    dc = o.DataCategory;

    sb.Append(o.Name + '\t' + p.Name + '\t' + o.Source + '\t' + m + '\t' + dc + '\t' + o.Description);
    sb.Append(Environment.NewLine);
    }
}

System.IO.File.WriteAllText(folderName + @"\Tables.txt", sb.ToString());

/**********************************************************************/
/*************************MEASURES & COLUMNS***************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("ObjectName" + '\t' + "TableName" + '\t' + "ObjectType" + '\t' + 
          "SourceColumn" + '\t' + "DataType" + '\t' + "Expression" + '\t' +
          "HiddenFlag" + '\t' + "Format" + '\t' + "PrimaryKey" + '\t' + "SummarizeBy" +'\t' + "DisplayFolder" + '\t' + "DataCategory" + '\t' + "SortByColumn" + '\t' + "Description");
sb.Append(Environment.NewLine);

// Columns
foreach (var t in Model.Tables.ToList())
{
    foreach (var o in t.Columns.ToList())
    {
        var sc = string.Empty;
        var dt = o.DataType;
        var expr = string.Empty;
        var hid = string.Empty;
        var fmt = o.FormatString;
        var pk = string.Empty;
        var sumb = o.SummarizeBy.ToString();
        var sbc = string.Empty;
        var fs = o.FormatString;
        
        if (o.Type == ColumnType.Data)
        {
            sc = (Model.Tables[t.Name].Columns[o.Name] as DataColumn).SourceColumn;      
        }
        else if (o.Type == ColumnType.Calculated)
        {
            expr = (Model.Tables[t.Name].Columns[o.Name] as CalculatedColumn).Expression;
            // Remove tabs and new lines
            expr = expr.Replace("\n"," ");
            expr = expr.Replace("\t"," ");
        }
        
        if (o.IsHidden == true)
        {
            hid = "Yes";
        }
        else
        {
            hid = "No";
        }
        
        if (o.IsKey == true)
        {
            pk = "Yes";
        }
        else
        {
            pk = "No";
        }
        
        try 
        {
            sbc = o.SortByColumn.Name;
        }
        catch
        {
            // Do nothing
        }

        if (sumb == "Default")
        {
            sumb = "";
        }
        
        sb.Append(o.Name + '\t' + t.Name + '\t' + "Column" + '\t' + sc + '\t' + dt + '\t' + expr + '\t' + 
        hid + '\t' + fs + '\t' + pk + '\t' + sumb + '\t' + o.DisplayFolder + '\t' + o.DataCategory + '\t' + 
        sbc + '\t' + o.Description);
        sb.Append(Environment.NewLine);
    }
}

foreach (var t in Model.Tables.ToList())
{
    foreach (var o in t.Measures.ToList())
    {
        var hid = string.Empty;
        var fs = o.FormatString;
        
        if (o.IsHidden == true)
        {
            hid = "Yes";
        }
        else
        {
            hid = "No";
        }
        
        var expr = o.Expression;
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ");
        expr = expr.Replace("\t"," ");
        
        sb.Append(o.Name + '\t' + t.Name + '\t' + "Measure" + '\t' + "" + '\t' + "" + '\t' + expr + '\t' + hid + '\t' + 
        fs + '\t' + "" + '\t' + "" + '\t' + o.DisplayFolder + '\t' + "" + '\t' + "" + '\t' + o.Description);
        sb.Append(Environment.NewLine);
    }
}

System.IO.File.WriteAllText(folderName + @"\MeasuresColumns.txt", sb.ToString());

/**********************************************************************/
/***********************************MODEL******************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("ModelName" + '\t' + "DefaultMode" + '\t' + "PowerBIDataSourceVersion");
sb.Append(Environment.NewLine);

var dm = "Import";
var pbi = string.Empty;

if (Model.DefaultMode == ModeType.DirectQuery)
{
    dm = "DirectQuery";
}

if (Model.DefaultPowerBIDataSourceVersion == PowerBIDataSourceVersion.PowerBI_V3)
{
    pbi = "Yes";
}

sb.Append(Model.Database.Name + '\t' + dm + '\t' + pbi);

System.IO.File.WriteAllText(folderName + @"\Model.txt", sb.ToString());

/**********************************************************************/
/***********************************ROLES******************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("RoleName" + '\t' + "RoleMember" + '\t' + "ModelPermission");
sb.Append(Environment.NewLine);

foreach (var r in Model.Roles.ToList())
{
    foreach (var rm in r.Members.ToList())
    {
        var mp = string.Empty;
        if (r.ModelPermission == ModelPermission.Administrator)
        {
            mp = r.ModelPermission.ToString().Substring(0,5);
        }
        else
        {
            mp = r.ModelPermission.ToString();
        }
        sb.Append(r.Name + '\t' + rm.Name + '\t' + mp);
        sb.Append(Environment.NewLine);
    }
}

System.IO.File.WriteAllText(folderName + @"\Roles.txt", sb.ToString());

/**********************************************************************/
/***********************************RLS********************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("RoleName" + '\t' + "TableName" + '\t' + "FilterExpression");
sb.Append(Environment.NewLine);

foreach (var t in Model.Tables)
{
    foreach(var r in Model.Roles)
    {
        var rls = Model.Tables[t.Name].RowLevelSecurity[r.Name];
        if (!String.IsNullOrEmpty(rls))
        {
            sb.Append(r.Name + '\t' + t.Name + '\t' + rls);
            sb.Append(Environment.NewLine);
        }
    }
}

System.IO.File.WriteAllText(folderName + @"\RLS.txt", sb.ToString());

/**********************************************************************/
/****************************RELATIONSHIPS*****************************/

sb = new System.Text.StringBuilder();

// Header
sb.Append("FromTable" + '\t' + "FromColumn" + '\t' + "ToTable" + '\t' + "ToColumn" + '\t' +
          "Active" + '\t' + "CrossFilteringBehavior" + '\t' + "RelationshipType" + '\t' + "SecurityFilteringBehavior" + '\t' + "RelyOnReferentialIntegrity");
sb.Append(Environment.NewLine);

foreach (var r in Model.Relationships)
{
    var act = string.Empty;
    var relType = string.Empty;
    var cfb = string.Empty;
    var sfb = string.Empty;
    var rori = string.Empty;
    
    if (r.IsActive == true)
    {
        act = "Yes";
    }
    else
    {
        act = "No";
    }
    
    if (r.FromCardinality == RelationshipEndCardinality.Many && r.ToCardinality ==  RelationshipEndCardinality.One)
    {
        relType = "Many-to-One";
    }
    else if (r.FromCardinality == RelationshipEndCardinality.Many && r.ToCardinality ==  RelationshipEndCardinality.Many)
    {
        relType = "Many-to-Many";
    }
    
    if (r.CrossFilteringBehavior == CrossFilteringBehavior.OneDirection)
    {
        cfb = "Single";
    }
    else if (r.CrossFilteringBehavior == CrossFilteringBehavior.BothDirections)
    {
        cfb = "Bi";
    }
    if (r.SecurityFilteringBehavior == SecurityFilteringBehavior.OneDirection)
    {
        sfb = "Single";
    }
    else if (r.SecurityFilteringBehavior == SecurityFilteringBehavior.BothDirections)
    {
        sfb = "Bi";
    }
    if (r.RelyOnReferentialIntegrity == true)
    {
        rori = "Yes";
    }
    else
    {
        rori = "No";
    }
    
    sb.Append(r.FromTable.Name + '\t' + r.FromColumn.Name + '\t' + r.ToTable.Name + '\t' + r.ToColumn.Name + '\t' +
    act + '\t' + cfb + '\t' + relType + '\t' + sfb + '\t' + rori);
    sb.Append(Environment.NewLine);
}

System.IO.File.WriteAllText(folderName + @"\Relationships.txt", sb.ToString());

/**********************************************************************/
/******************************HIERARCHIES*****************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("HierarchyName" + '\t' + "TableName" + '\t' + "ColumnName");
sb.Append(Environment.NewLine);

// Hierarchies
foreach (var h in Model.AllHierarchies.ToList())
{
    foreach (var lev in h.Levels.ToList())
    {
        sb.Append(h.Name + '\t' + h.Table.Name + '\t' + lev.Name);
        sb.Append(Environment.NewLine);
    }
}

System.IO.File.WriteAllText(folderName + @"\Hierarchies.txt", sb.ToString());

/**********************************************************************/
/******************************PERSPECTIVES****************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("TableName" + '\t' + "ObjectName" + '\t' + "ObjectType");

// Loop header for each perspective
foreach (var p in Model.Perspectives.ToList())
{
    sb.Append('\t' + p.Name);
}

sb.Append(Environment.NewLine);

// Measures
foreach (var o in Model.AllMeasures.ToList())
{
    sb.Append(o.Parent.Name + '\t' + o.Name + '\t' + "Measure");
    
    foreach (var p in Model.Perspectives.ToList())
    {
        var per = string.Empty;
        if (o.InPerspective[p] == true)
        {
            per = "Yes";
        }
        else
        {
            per = "No";
        }
        sb.Append('\t' + per);
    }
    sb.Append(Environment.NewLine);
}

// Columns
foreach (var o in Model.AllColumns.ToList())
{
    sb.Append(o.Table.Name + '\t' + o.Name + '\t' + "Column");
    
    foreach (var p in Model.Perspectives.ToList())
    {
        var per = string.Empty;
        if (o.InPerspective[p] == true)
        {
            per = "Yes";
        }
        else
        {
            per = "No";
        }
        sb.Append('\t' + per);
    }
    sb.Append(Environment.NewLine);
}

// Hierarchies
foreach (var o in Model.AllHierarchies.ToList())
{
    sb.Append(o.Parent.Name + '\t' + o.Name + '\t' + "Hierarchy");
    
    foreach (var p in Model.Perspectives.ToList())
    {
        var per = string.Empty;
        if (o.InPerspective[p] == true)
        {
            per = "Yes";
        }
        else
        {
            per = "No";
        }
        sb.Append('\t' + per);
    }
    sb.Append(Environment.NewLine);
}

System.IO.File.WriteAllText(folderName + @"\Perspectives.txt", sb.ToString());

/**********************************************************************/
/******************************TRANSLATIONS****************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("ObjectName" + '\t' + "ObjectType" + '\t' + "TableName" + '\t' + "TranslationLanguage" + '\t' + "TranslatedObjectName" + '\t' + "TranslatedObjectDescription" + '\t' + "TranslatedDisplayFolder");
sb.Append(Environment.NewLine);


// Add placeholder if no translations exist
bool hasTranslation = true;
var placeLang = "00-00";

if (Model.Cultures.Count() == 0)
{
    hasTranslation = false;
    Model.AddTranslation(placeLang);
}
    
foreach (var cul in Model.Cultures.ToList()) 
{ 
    // Tables
    foreach (var t in Model.Tables.ToList())
    {
    
        var objectName = t.Name;
        var objectType = "Table";
        var tableName = t.Name;
        var transLang = string.Empty;
        var transName = string.Empty;
        var transDesc = string.Empty;
        var transDF = string.Empty;
    
        if (hasTranslation == true)
        {
            transLang = cul.Name;
            transName = t.TranslatedNames[transLang];
            transDesc = t.TranslatedDescriptions[transLang];
            transDF = string.Empty;
        }
        
        sb.Append(objectName + '\t' + objectType + '\t' + tableName + '\t' + transLang + '\t' + transName + '\t' + transDesc + '\t' + transDF);
        sb.Append(Environment.NewLine);
    }
    
    // Columns
    foreach (var c in Model.AllColumns.ToList())
    {
        var objectName = c.Name;
        var objectType = "Column";
        var tableName = c.Table.Name;
        var transLang = string.Empty;
        var transName = string.Empty;
        var transDesc = string.Empty;
        var transDF = string.Empty;
                
        if (hasTranslation == true)
        {
            transLang = cul.Name;
            transName = c.TranslatedNames[transLang];
            transDesc = c.TranslatedDescriptions[transLang];
            transDF = c.TranslatedDisplayFolders[transLang];
        }
        
        sb.Append(objectName + '\t' + objectType + '\t' + tableName + '\t' + transLang + '\t' + transName + '\t' + transDesc + '\t' + transDF);
        sb.Append(Environment.NewLine);
    }
    
    // Measures
    foreach (var m in Model.AllMeasures.ToList())
    {
        var objectName = m.Name;
        var objectType = "Measure";
        var tableName = m.Table.Name;
        var transLang = string.Empty;
        var transName = string.Empty;
        var transDesc = string.Empty;
        var transDF = string.Empty;
        
        if (hasTranslation == true)
        {
            transLang = cul.Name;
            transName = m.TranslatedNames[transLang];
            transDesc = m.TranslatedDescriptions[transLang];
            transDF = m.TranslatedDisplayFolders[transLang];
        }
        
        sb.Append(objectName + '\t' + objectType + '\t' + tableName + '\t' + transLang + '\t' + transName + '\t' + transDesc + '\t' + transDF);
        sb.Append(Environment.NewLine);
    }
    
    // Hierarchies
    foreach (var h in Model.AllHierarchies.ToList())
    {
        var objectName = h.Name;
        var objectType = "Hierarchy";
        var tableName = h.Table.Name;
        var transLang = string.Empty;
        var transName = string.Empty;
        var transDesc = string.Empty;
        var transDF = string.Empty;
        
        if (hasTranslation == true)
        {
            transLang = cul.Name;
            transName = h.TranslatedNames[transLang];
            transDesc = h.TranslatedDescriptions[transLang];
            transDF = h.TranslatedDisplayFolders[transLang];
        }
        
        sb.Append(objectName + '\t' + objectType + '\t' + tableName + '\t' + transLang + '\t' + transName + '\t' + transDesc + '\t' + transDF);
        sb.Append(Environment.NewLine);
    }
}

// Remove the placeholder translation
if (hasTranslation == false)
{
    Model.Cultures[placeLang].Delete();
}

System.IO.File.WriteAllText(folderName + @"\Translations.txt", sb.ToString());

/**********************************************************************/
/*************************CALCULATION GROUPS***************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("CalculationGroup" + '\t' + "CalculationItem" + '\t' + "Expression" + '\t' + "Ordinal" + '\t' + "FormatString" + '\t' + "Description");
sb.Append(Environment.NewLine);

foreach (var o in Model.CalculationGroups.ToList())
{
    foreach (var i in o.CalculationItems.ToList())
    {
        var cg = o.Name;
        var ci = i.Name;
        var expr = i.Expression;
        
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ");
        expr = expr.Replace("\t"," ");

        var ord = i.Ordinal.ToString();
        var fs = i.FormatStringExpression;
        var desc = i.Description;
        
        
        sb.Append(cg + '\t' + ci + '\t' + expr + '\t' + ord + '\t' + fs + '\t' + desc);
        sb.Append(Environment.NewLine);
    }

}

System.IO.File.WriteAllText(folderName + @"\CalculationGroups.txt", sb.ToString());