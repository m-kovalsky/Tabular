# Tabular
This repo is a collection of useful code for automating processes within tabular modeling. All of these scripts are to be executed in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") so make sure to download and install it.

For addtional information on these scripts and more, check out my blog [Elegant BI](https://www.elegantbi.com "Elegant BI").

### [Automated Data Dictionary](https://www.elegantbi.com/post/datadictionaryreinvented "Automated Data Dictionary")

Run this script in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") to create an automated data dictionary. This script works for Analysis Services, Azure Analysis Services, and Power BI Premium models ([XMLA R/W endpoints](https://docs.microsoft.com/en-us/power-bi/admin/service-premium-connect-tools#enable-xmla-read-write "XMLA R/W endpoints") enabled).

### [Automated Data Dictionary via Excel](https://www.elegantbi.com/post/datadictionaryexcel "Automated Data Dictionary via Excel")

Run this script in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") to create an automated data dictionary where the Data Dictionary table is stored in Excel. This script works for Analysis Services, Azure Analysis Services, and Power BI Premium models ([XMLA R/W endpoints](https://docs.microsoft.com/en-us/power-bi/admin/service-premium-connect-tools#enable-xmla-read-write "XMLA R/W endpoints") enabled).

### [Blank Row Finder](https://www.elegantbi.com/post/findblankrows "Blank Row Finder")

Run this script in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") against a live-connected model to quickly make a list of all relationships that contain a blank row in the 'to-table'.

### [Descriptions](https://github.com/m-kovalsky/Tabular/tree/master/Descriptions "Descriptions")

Run the ExportDescriptions.cs script in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") to export objects and existing descriptions in your tabular model to an Excel file.

Run the ImportDescriptions.cs script in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") to import object descriptions back into your tabular model from the Excel file.

### [Export BPA Results](https://www.elegantbi.com/post/exportbparesults "Export BPA Results")

Running this script in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") will run the [Best Practice Analyzer](https://docs.tabulareditor.com/Best-Practice-Analyzer.html "Best Practice Analyzer") and output the results. The output can easily be copied into Excel for further analysis.

### [Cancel Processing](https://www.elegantbi.com/post/canceldatarefreshte "Cancel Processing")

Run this script in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") against a live-connected model to cancel the data refresh of that model.

### [Master Model](https://www.elegantbi.com/post/mastermodel "Master Model")

### [Metadata Export](https://www.elegantbi.com/post/extractmodelmetadata "Metadata Export")

### Metadata Import - Perspectives
Run this script to automatically update the perspectives in your model (or add new perspectives). This script coordinates with the output text file from the Metadata Export script.

### Metadata Import - Translations
Run this script to automatically update the translations in your model (or add new translations). This script coordinates with the output text file from the Metadata Export script.

### [Auto Aggs](https://www.elegantbi.com/post/autoaggs "Auto Aggs")

### [Perspective Editor](https://www.elegantbi.com/post/perspectiveeditor "Perspective Editor")

Running this script opens a program within [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") that allows you to create or modify perspectives akin to the way it is done in SQL Server Development Tools (SSDT). It also gives you a tree-view of all the objects that are in a perspective relative to all the objects in the model.

### [Report Level Measures](https://www.elegantbi.com/post/reportlevelmeasures "Report Level Measures")

Want to migrate measures created within a Power BI Desktop report to your tabular model? This script does exactly that. Setting the 'createMeasures' parameter to 'true' will create the measures in the model file within Tabular Editor. Setting this paramter to 'false' will dynamically generate C# code which can be copied and executed in order to create the measures in a model.

### [Vertipaq Annotations](https://www.elegantbi.com/post/vertipaqintabulareditor "Vertipaq Annotations")

Run this script against a live-connected model to save [Vertipaq Analyzer](https://www.sqlbi.com/tools/vertipaq-analyzer/ "Vertipaq Analyzer") statistics as annotations on model objects. These annotations may be referenced to create Best Practice Analyzer rules for your model. See the link below for more info on [Tabular Editor](https://tabulareditor.com/ "Tabular Editor")'s [Best Practice Analyzer](https://docs.tabulareditor.com/Best-Practice-Analyzer.html "Best Practice Analyzer").

*Note: If running this script against a Power BI Desktop model (using Tabular Editor as an External Tool), you must select the following setting within Tabular Editor:*

    File -> Preferences -> Features -> Allow Unsupported Power BI features (experimental)

* **Model:** Model Size

* **Tables:** Row Count; Table Size; Table Size as a Percentage of the Model Size

* **Partitions:** Record Count; Segment Count; Records Per Segment

* **Columns:** Cardinality; Column Hierarchy Size; Column Size; Data Size; Dictionary Size; Column Size as a Percentage of the Table Size; Column Size as a Percentage of the Model Size

* **Hierarchies:** User Hierarchy Size

* **Relationships:** Relationship Size; Max From Cardinality; Max To Cardinality

### [Vpax to Tabular Editor](https://www.elegantbi.com/post/vpaxtotabulareditor "Vpax to Tabular Editor")

This script creates the same annotations as the [Vertipaq Annotations](https://github.com/m-kovalsky/Tabular/blob/master/VertipaqAnnotations.cs "Vertipaq Annotations") script. The only difference is that this script loads the [Vertipaq Analyzer](https://www.sqlbi.com/tools/vertipaq-analyzer/ "Vertipaq Analyzer") data from a Vertipaq Analyzer (.vpax) file. The .vpax file can be generated by selecting 'View Metrics' within the 'Advanced' tab in [DAX Studio](https://daxstudio.org/ "DAX Studio").
