var folderName = @"C:\Users\mikova\Desktop\Metadata"; //Update this to the folder that contains the Metadata Export text files

var fileName = @"\Translations.txt";

var Metadata = ReadFile(folderName+fileName);

// Split the file into rows by CR and LF characters:
var tsvRows = Metadata.Split(new[] {'\r','\n'},StringSplitOptions.RemoveEmptyEntries);

// Loop through all rows but skip the first one:
foreach(var row in tsvRows.Skip(1))
{
    var tsvColumns = row.Split('\t');     // Assume file uses tabs as column separator  
    var objectName = tsvColumns[0];
    var objectType = tsvColumns[1];
    var tableName = tsvColumns[2];
    var lang = tsvColumns[3];
    var transobj = tsvColumns[4];
    var transobjdesc = tsvColumns[5];
    var transobjdf = tsvColumns[6];
    
    // Add new translation language if it does not already exist
    if (Model.Cultures.Where(a => a.Name == lang).Count() == 0)
    {
        Model.AddTranslation(lang);
    }
    
    // Add the translations for tables
    if (objectType == "Table")
    {
        Model.Tables[tableName].TranslatedNames[lang] = transobj;
        Model.Tables[tableName].TranslatedDescriptions[lang] = transobjdesc;
    }

    // Add the translations for columns
    else if (objectType == "Column")
    {
        Model.Tables[tableName].Columns[objectName].TranslatedNames[lang] = transobj;
        Model.Tables[tableName].Columns[objectName].TranslatedDescriptions[lang] = transobjdesc;
        Model.Tables[tableName].Columns[objectName].TranslatedDisplayFolders[lang] = transobjdf;
        
    }

    // Add the translations for measures
    else if (objectType == "Measure")
    {
        Model.Tables[tableName].Measures[objectName].TranslatedNames[lang] = transobj;
        Model.Tables[tableName].Measures[objectName].TranslatedDescriptions[lang] = transobjdesc;
        Model.Tables[tableName].Measures[objectName].TranslatedDisplayFolders[lang] = transobjdf;
    }
    
    // Add the translations for hierarchies
    else if (objectType == "Hierarchy")
    {
        Model.Tables[tableName].Hierarchies[objectName].TranslatedNames[lang] = transobj;
        Model.Tables[tableName].Hierarchies[objectName].TranslatedDescriptions[lang] = transobjdesc;
        Model.Tables[tableName].Hierarchies[objectName].TranslatedDisplayFolders[lang] = transobjdf;
    }
}