string folderName = @"C:\Desktop\Metadata"; // Update this location to the destination folder on your computer

var sb = new System.Text.StringBuilder();
string newline = Environment.NewLine;

/******************************DATA SOURCES****************************/

// Headers
sb.Append("DataSource" + '\t' + "ConnectionString" + '\t' + "Provider" + '\t' + "MaxConnections" + newline);

foreach (var o in Model.DataSources.Where(a => a.Type.ToString() == "Provider").OrderBy(a => a.Name).ToList())
{
    var ds = (Model.DataSources[o.Name] as ProviderDataSource);
    string n = o.Name;
    string conn = ds.ConnectionString;
    string mc = ds.MaxConnections.ToString();
    string prov = ds.Provider;
    
    sb.Append(n + '\t' + conn + '\t' + prov + '\t' + mc + newline);
}

System.IO.File.WriteAllText(folderName + @"\DataSources.txt", sb.ToString());

/**********************************************************************/
/*********************************TABLES*******************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("TableName" + '\t' + "PartitionName" + '\t' + "DataSource" + '\t' + "Mode" + '\t' +
          "DataCategory" + '\t' + "Description" + '\t' + "Query" + newline);
          
foreach (var o in Model.Tables.Where(a => a.ObjectTypeName != "Calculation Group Table").OrderBy(a => a.Name).ToList())
{
    foreach (var p in o.Partitions.OrderBy(a => a.Name).ToList())
    {
        string q = Model.Tables[o.Name].Partitions[p.Name].Query.Replace("\t"," ").Replace("\r"," ").Replace("\n"," ");
        string m = Model.Tables[o.Name].Partitions[p.Name].Mode.ToString();
        string dc = o.DataCategory;

        sb.Append(o.Name + '\t' + p.Name + '\t' + o.Source + '\t' + m + '\t' + dc + '\t' + o.Description + '\t' + q + newline);
    }
}

System.IO.File.WriteAllText(folderName + @"\Tables.txt", sb.ToString());

/**********************************************************************/
/*************************MEASURES & COLUMNS***************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("ObjectName" + '\t' + "TableName" + '\t' + "ObjectType" + '\t' + "SourceColumn" + '\t' + "DataType" + '\t' + "Expression" + '\t' +
          "HiddenFlag" + '\t' + "Format" + '\t' + "PrimaryKey" + '\t' + "SummarizeBy" +'\t' + "DisplayFolder" + '\t' + "DataCategory" + '\t' + "SortByColumn" + '\t' + "Description" + '\t' + "EncodingHint" + newline);

// Columns
foreach (var t in Model.Tables.OrderBy(a => a.Name).ToList())
{
    foreach (var o in t.Columns.OrderBy(a => a.Name).ToList())
    {
        string sc = string.Empty;
        string dt = o.DataType.ToString();
        string expr = string.Empty;
        string hid = string.Empty;
        string pk = string.Empty;
        string sumb = o.SummarizeBy.ToString();
        string sbc = string.Empty;
        string fs = o.FormatString;
        string eh = o.EncodingHint.ToString();
        
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
        sbc + '\t' + o.Description + '\t' + eh + newline);
    }
}

foreach (var t in Model.Tables.OrderBy(a => a.Name).ToList())
{
    foreach (var o in t.Measures.OrderBy(a => a.Name).ToList())
    {
        string hid = string.Empty;
        string fs = o.FormatString;
        string expr = o.Expression;
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ");
        expr = expr.Replace("\t"," ");
        
        if (o.IsHidden)
        {
            hid = "Yes";
        }
        else
        {
            hid = "No";
        }
        
        sb.Append(o.Name + '\t' + t.Name + '\t' + "Measure" + '\t' + "" + '\t' + "" + '\t' + expr + '\t' + hid + '\t' + 
        fs + '\t' + "" + '\t' + "" + '\t' + o.DisplayFolder + '\t' + "" + '\t' + "" + '\t' + o.Description + "N/A" +newline);
    }
}

System.IO.File.WriteAllText(folderName + @"\MeasuresColumns.txt", sb.ToString());

/**********************************************************************/
/***********************************MODEL******************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("ModelName" + '\t' + "DefaultMode" + '\t' + "PowerBIDataSourceVersion" + newline);

string dm = "Import";
string pbi = string.Empty;

if (Model.DefaultMode == ModeType.DirectQuery)
{
    dm = "DirectQuery";
}

if (Model.DefaultPowerBIDataSourceVersion == PowerBIDataSourceVersion.PowerBI_V3)
{
    pbi = "Yes";
}

sb.Append(Model.Database.Name + '\t' + dm + '\t' + pbi + newline);

System.IO.File.WriteAllText(folderName + @"\Model.txt", sb.ToString());

/**********************************************************************/
/***********************************ROLES******************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("RoleName" + '\t' + "RoleMember" + '\t' + "ModelPermission" + newline);

foreach (var r in Model.Roles.OrderBy(a => a.Name).ToList())
{
    foreach (var rm in r.Members.OrderBy(a => a.Name).ToList())
    {
        string mp = string.Empty;
        if (r.ModelPermission == ModelPermission.Administrator)
        {
            mp = r.ModelPermission.ToString().Substring(0,5);
        }
        else
        {
            mp = r.ModelPermission.ToString();
        }
        sb.Append(r.Name + '\t' + rm.Name + '\t' + mp + newline);    
    }
}

System.IO.File.WriteAllText(folderName + @"\Roles.txt", sb.ToString());

/**********************************************************************/
/***********************************RLS********************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("RoleName" + '\t' + "TableName" + '\t' + "FilterExpression" + newline);

foreach (var r in Model.Roles.OrderBy(a => a.Name).ToList())
{
    foreach(var t in Model.Tables.OrderBy(a => a.Name).ToList())
    {
        string rls = Model.Tables[t.Name].RowLevelSecurity[r.Name];
        if (!String.IsNullOrEmpty(rls))
        {
            sb.Append(r.Name + '\t' + t.Name + '\t' + rls + newline);
        }
    }
}

System.IO.File.WriteAllText(folderName + @"\RLS.txt", sb.ToString());

/**********************************************************************/
/****************************RELATIONSHIPS*****************************/

sb = new System.Text.StringBuilder();

// Header
sb.Append("FromTable" + '\t' + "FromColumn" + '\t' + "ToTable" + '\t' + "ToColumn" + '\t' +
          "Active" + '\t' + "CrossFilteringBehavior" + '\t' + "RelationshipType" + '\t' + "SecurityFilteringBehavior" + '\t' + "RelyOnReferentialIntegrity" + newline);

foreach (var r in Model.Relationships.ToList())
{
    string actv = string.Empty;
    string relType = string.Empty;
    string cfb = string.Empty;
    string sfb = string.Empty;
    string rori = string.Empty;
    
    if (r.IsActive)
    {
        actv = "Yes";
    }
    else
    {
        actv = "No";
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
    if (r.RelyOnReferentialIntegrity)
    {
        rori = "Yes";
    }
    else
    {
        rori = "No";
    }
    
    sb.Append(r.FromTable.Name + '\t' + r.FromColumn.Name + '\t' + r.ToTable.Name + '\t' + r.ToColumn.Name + '\t' +
    actv + '\t' + cfb + '\t' + relType + '\t' + sfb + '\t' + rori + newline);
}

System.IO.File.WriteAllText(folderName + @"\Relationships.txt", sb.ToString());

/**********************************************************************/
/******************************HIERARCHIES*****************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("HierarchyName" + '\t' + "TableName" + '\t' + "ColumnName" + '\t' + "Ordinal" + newline);

// Hierarchies
foreach (var h in Model.AllHierarchies.OrderBy(a => a.Name).ToList())
{
    foreach (var lev in h.Levels.OrderBy(a => a.Name).ToList())
    {
        sb.Append(h.Name + '\t' + h.Table.Name + '\t' + lev.Name + '\t' + lev.Ordinal.ToString() + newline);
    }
}

System.IO.File.WriteAllText(folderName + @"\Hierarchies.txt", sb.ToString());

/**********************************************************************/
/******************************PERSPECTIVES****************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("TableName" + '\t' + "ObjectName" + '\t' + "ObjectType");

// Loop header for each perspective
foreach (var p in Model.Perspectives.OrderBy(a => a.Name).ToList())
{
    sb.Append('\t' + p.Name);
}

sb.Append(newline);

// Measures
foreach (var o in Model.AllMeasures.OrderBy(a => a.Name).ToList())
{
    sb.Append(o.Parent.Name + '\t' + o.Name + '\t' + "Measure");
    
    foreach (var p in Model.Perspectives.OrderBy(a => a.Name).ToList())
    {
        string per = string.Empty;
        if (o.InPerspective[p])
        {
            per = "Yes";
        }
        else
        {
            per = "No";
        }
        sb.Append('\t' + per);
    }
    sb.Append(newline);
}

// Columns
foreach (var o in Model.AllColumns.OrderBy(a => a.Name).ToList())
{
    sb.Append(o.Table.Name + '\t' + o.Name + '\t' + "Column");
    
    foreach (var p in Model.Perspectives.OrderBy(a => a.Name).ToList())
    {
        string per = string.Empty;
        if (o.InPerspective[p])
        {
            per = "Yes";
        }
        else
        {
            per = "No";
        }
        sb.Append('\t' + per);
    }
    sb.Append(newline);
}

// Hierarchies
foreach (var o in Model.AllHierarchies.OrderBy(a => a.Name).ToList())
{
    sb.Append(o.Parent.Name + '\t' + o.Name + '\t' + "Hierarchy");
    
    foreach (var p in Model.Perspectives.OrderBy(a => a.Name).ToList())
    {
        string per = string.Empty;
        if (o.InPerspective[p])
        {
            per = "Yes";
        }
        else
        {
            per = "No";
        }
        sb.Append('\t' + per);
    }
    sb.Append(newline);
}

System.IO.File.WriteAllText(folderName + @"\Perspectives.txt", sb.ToString());

/**********************************************************************/
/******************************TRANSLATIONS****************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("ObjectName" + '\t' + "ObjectType" + '\t' + "TableName" + '\t' + "TranslationLanguage" + '\t' + "TranslatedObjectName" + '\t' + "TranslatedObjectDescription" + '\t' + "TranslatedDisplayFolder" + newline);

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
    foreach (var t in Model.Tables.OrderBy(a => a.Name).ToList())
    {    
        string objectName = t.Name;
        string objectType = "Table";
        string tableName = t.Name;
        string transLang = string.Empty;
        string transName = string.Empty;
        string transDesc = string.Empty;
        string transDF = string.Empty;
    
        if (hasTranslation)
        {
            transLang = cul.Name;
            transName = t.TranslatedNames[transLang];
            transDesc = t.TranslatedDescriptions[transLang];
            transDF = string.Empty;
        }
        
        sb.Append(objectName + '\t' + objectType + '\t' + tableName + '\t' + transLang + '\t' + transName + '\t' + transDesc + '\t' + transDF + newline);
    }
    
    // Columns
    foreach (var c in Model.AllColumns.OrderBy(a => a.Name).ToList())
    {
        string objectName = c.Name;
        string objectType = "Column";
        string tableName = c.Table.Name;
        string transLang = string.Empty;
        string transName = string.Empty;
        string transDesc = string.Empty;
        string transDF = string.Empty;
                
        if (hasTranslation)
        {
            transLang = cul.Name;
            transName = c.TranslatedNames[transLang];
            transDesc = c.TranslatedDescriptions[transLang];
            transDF = c.TranslatedDisplayFolders[transLang];
        }
        
        sb.Append(objectName + '\t' + objectType + '\t' + tableName + '\t' + transLang + '\t' + transName + '\t' + transDesc + '\t' + transDF + newline);
    }
    
    // Measures
    foreach (var m in Model.AllMeasures.OrderBy(a => a.Name).ToList())
    {
        string objectName = m.Name;
        string objectType = "Measure";
        string tableName = m.Table.Name;
        string transLang = string.Empty;
        string transName = string.Empty;
        string transDesc = string.Empty;
        string transDF = string.Empty;
        
        if (hasTranslation)
        {
            transLang = cul.Name;
            transName = m.TranslatedNames[transLang];
            transDesc = m.TranslatedDescriptions[transLang];
            transDF = m.TranslatedDisplayFolders[transLang];
        }
        
        sb.Append(objectName + '\t' + objectType + '\t' + tableName + '\t' + transLang + '\t' + transName + '\t' + transDesc + '\t' + transDF + newline);
    }
    
    // Hierarchies
    foreach (var h in Model.AllHierarchies.OrderBy(a => a.Name).ToList())
    {
        string objectName = h.Name;
        string objectType = "Hierarchy";
        string tableName = h.Table.Name;
        string transLang = string.Empty;
        string transName = string.Empty;
        string transDesc = string.Empty;
        string transDF = string.Empty;
        
        if (hasTranslation)
        {
            transLang = cul.Name;
            transName = h.TranslatedNames[transLang];
            transDesc = h.TranslatedDescriptions[transLang];
            transDF = h.TranslatedDisplayFolders[transLang];
        }
        
        sb.Append(objectName + '\t' + objectType + '\t' + tableName + '\t' + transLang + '\t' + transName + '\t' + transDesc + '\t' + transDF + newline);
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
sb.Append("CalculationGroup" + '\t' + "CalculationItem" + '\t' + "Expression" + '\t' + "Ordinal" + '\t' + "FormatString" + '\t' + "Description" + newline);

foreach (var o in Model.CalculationGroups.OrderBy(a => a.Name).ToList())
{
    foreach (var i in o.CalculationItems.OrderBy(a => a.Name).ToList())
    {
        string cg = o.Name;
        string ci = i.Name;
        string expr = i.Expression;
        
        // Remove tabs and new lines
        expr = expr.Replace("\n"," ");
        expr = expr.Replace("\t"," ");

        string ord = i.Ordinal.ToString();
        string fs = i.FormatStringExpression;
        string desc = i.Description;
        
        sb.Append(cg + '\t' + ci + '\t' + expr + '\t' + ord + '\t' + fs + '\t' + desc + newline);
    }
}

System.IO.File.WriteAllText(folderName + @"\CalculationGroups.txt", sb.ToString());

/**********************************************************************/
/*********************************KPI**********************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("TableName" + '\t' + "MeasureName" + '\t' + "StatusExpresssion" + '\t' + "StatusGraphic" + '\t' +
          "StatusDescription" + '\t' + "TargetExpression" + '\t' + "TargetFormatString" + '\t' + "TargetDescription" + '\t' + "TrendExpression" + '\t' + "TrendGraphic" + '\t' + "TrendDescription" + newline);

foreach (var m in Model.AllMeasures.Where(a => a.KPI != null).OrderBy(a => a.Name).ToList())
{
    
    string tableName = m.Table.Name;
    string measureName = m.Name;
    var k = m.KPI;
    string statusExpr = k.StatusExpression.Replace("\n"," ").Replace("\t"," ");;
    string statusDesc = k.StatusDescription;
    string statusGraphic = k.StatusGraphic;
    string targetExpr = k.TargetExpression.Replace("\n"," ").Replace("\t"," ");
    string targetDesc = k.TargetDescription;
    string targetFS = k.TargetFormatString;
    string trendExpr = k.TrendExpression.Replace("\n"," ").Replace("\t"," ");;
    string trendDesc = k.TrendDescription;
    string trendGraphic = k.TrendGraphic;
 
    sb.Append(tableName + '\t' + measureName + '\t' + statusExpr + '\t' + statusGraphic + '\t' + statusDesc + '\t' + targetExpr + '\t' + targetFS + '\t' + targetDesc + '\t' + trendExpr + '\t' + trendGraphic + '\t' + trendDesc + newline);
}

System.IO.File.WriteAllText(folderName + @"\KPI.txt", sb.ToString());

/**********************************************************************/
/*********************************OLS**********************************/

sb = new System.Text.StringBuilder();

// Headers
sb.Append("RoleName" + '\t' + "ObjectType" + '\t' + "TableName" + '\t' + "ObjectName" + '\t' + "ObjectLevelSecurity" + newline);

if (Model.Database.CompatibilityLevel >= 1400)
{
    foreach (var t in Model.Tables.OrderBy(a => a.Name).ToList())
    {
        string tableName = t.Name;
        
        foreach(var r in Model.Roles.OrderBy(a => a.Name).ToList())
        {
            string roleName = r.Name;
            string tableOLS = Model.Tables[tableName].ObjectLevelSecurity[roleName].ToString();
            if (!String.IsNullOrEmpty(tableOLS))
            {
                sb.Append(roleName + '\t' + "Table" + '\t' + tableName + '\t' + tableName + '\t' + tableOLS + newline);
            }
            
            foreach (var c in t.Columns.OrderBy(a => a.Name).ToList())
            {
                string colName = c.Name;
                string colOLS = Model.Tables[tableName].Columns[colName].ObjectLevelSecurity[roleName].ToString();
                
                if (!String.IsNullOrEmpty(colOLS))
                {
                    sb.Append(roleName + '\t' + "Column" + '\t' + tableName + '\t' + colName + '\t' + colOLS + newline);
                }
            }
        }    
    }
}

System.IO.File.WriteAllText(folderName + @"\OLS.txt", sb.ToString());
