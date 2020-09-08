var folderName = @"C:\Users\mikova\Desktop\Metadata"; //Update this to the folder that contains the Metadata Export text files

var fileName = @"\Perspectives.txt";

var Metadata = ReadFile(folderName+fileName);

// Split the file into rows by CR and LF characters:
var tsvRows = Metadata.Split(new[] {'\r','\n'},StringSplitOptions.RemoveEmptyEntries);

// Set array for perspective names
var col = tsvRows[0].Split('\t');
int pCount = col.Count() - 3;
string[] p = new string[1000];

for (int i = 0; i < pCount; i++)
{
    p[i] = col[i+3];
}

// Loop through all rows but skip the first one:
foreach(var row in tsvRows.Skip(1))
{
    var tsvColumns = row.Split('\t');     // Assume file uses tabs as column separator  
    
    if (pCount > 0)
    {
        for (int i = 0; i < pCount; i++)
        {
            var tableName = tsvColumns[0];
            var objectName = tsvColumns[1];
            var objectType = tsvColumns[2];
            var persp =  p[i];
            bool yesno = true;
                
            if (tsvColumns[i+3] == "Yes")
            {
                yesno = true;
            }
            else if (tsvColumns[i+3] == "No")
            {
                yesno = false;
            }

            // Update perspective values for objects
            if (objectType == "Column")
            {
                Model.Tables[tableName].Columns[objectName].InPerspective[persp] = yesno;
            }
            
            else if (objectType == "Measure")
            {
                Model.Tables[tableName].Measures[objectName].InPerspective[persp] = yesno;
            }
         
            else if (objectType == "Hierarchy")
            {
                Model.Tables[tableName].Hierarchies[objectName].InPerspective[persp] = yesno;
            }    
        }
    }
}