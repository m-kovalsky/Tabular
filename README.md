# Tabular
Useful code for tabular modeling and automation

For more information on these scripts, check out my blog (https://www.elegantbi.com)

### Blank Row Finder
Run this script in Tabular Editor against a live-connected model to quickly make a list of all relationships that contain a blank row in the 'to-table'.
https://www.elegantbi.com/post/findblankrows

### Master Model
https://www.elegantbi.com/post/mastermodel

### Metadata Export
https://www.elegantbi.com/post/extractmodelmetadata

### Metadata Import - Perspectives
Run this script to automatically update the perspectives in your model (or add new perspectives). This script coordinates with the output text file from the Metadata Export script.

### Metadata Import - Translations
Run this script to automatically update the translations in your model (or add new translations). This script coordinates with the output text file from the Metadata Export script.

### Auto Aggs
https://www.elegantbi.com/post/autoaggs

### Vertipaq Annotations
https://www.elegantbi.com/post/vertipaqintabulareditor

Run this script against a live-connected model to save Vertipaq Analyzer statistics as annotations on model objects. These annotations may be referenced to create Best Practice Analyzer rules for your model. See the link below for more info on Tabular Editor's Best Practice Analyzer.

https://docs.tabulareditor.com/Best-Practice-Analyzer.html

* **Model:** Model Size

* **Tables:** Row Count; Table Size; Table Size as a Percentage of the Model Size

* **Partitions:** Record Count; Segment Count; Records Per Segment

* **Columns:** Cardinality; Column Hierarchy Size; Column Size; Data Size; Dictionary Size; Column Size as a Percentage of the Table Size; Column Size as a Percentage of the Model Size

* **Hierarchies:** User Hierarchy Size

* **Relationships:** Relationship Size; Max From Cardinality; Max To Cardinality

### Perspective Editor
Running this script opens a program within Tabular Editor that allows you to create or modify perspectives akin to the way it is done in SQL Server Development Tools (SSDT). It also gives you a tree-view of all the objects that are in a perspective relative to all the objects in the model.
