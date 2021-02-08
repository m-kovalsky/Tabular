# Tabular
Useful code for tabular modeling and automation

For more information on these scripts, check out my [blog](https://www.elegantbi.com "blog").

### [Automated Data Dictionary](https://www.elegantbi.com/post/datadictionaryreinvented "Automated Data Dictionary")

Run this script in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") to create an automated data dictionary. This script works for Analysis Services, Azure Analysis Services, and Power BI Premium models (XMLA R/W Endpoints enabled).

### [Blank Row Finder](https://www.elegantbi.com/post/findblankrows "Blank Row Finder")

Run this script in [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") against a live-connected model to quickly make a list of all relationships that contain a blank row in the 'to-table'.

### [Master Model](https://www.elegantbi.com/post/mastermodel "Master Model")

### [Metadata Export](https://www.elegantbi.com/post/extractmodelmetadata "Metadata Export")

### Metadata Import - Perspectives
Run this script to automatically update the perspectives in your model (or add new perspectives). This script coordinates with the output text file from the Metadata Export script.

### Metadata Import - Translations
Run this script to automatically update the translations in your model (or add new translations). This script coordinates with the output text file from the Metadata Export script.

### [Auto Aggs](https://www.elegantbi.com/post/autoaggs "Auto Aggs")

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

### [Perspective Editor](https://www.elegantbi.com/post/perspectiveeditor "Perspective Editor")

Running this script opens a program within [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") that allows you to create or modify perspectives akin to the way it is done in SQL Server Development Tools (SSDT). It also gives you a tree-view of all the objects that are in a perspective relative to all the objects in the model.

### [Vpax to Tabular Editor](https://www.elegantbi.com/post/vpaxtotabulareditor "Vpax to Tabular Editor")

This script creates the same annotations as the [Vertipaq Annotations](https://github.com/m-kovalsky/Tabular/blob/master/VertipaqAnnotations.cs "Vertipaq Annotations") script. The only difference is that this script loads the [Vertipaq Analyzer](https://www.sqlbi.com/tools/vertipaq-analyzer/ "Vertipaq Analyzer") data from a Vertipaq Analyzer (.vpax) file.
