# Description

This is FreeCAD macro to export spreadsheet to file. It has been created to support my woodworking project [getDimensions](https://github.com/dprojects/getDimensions). However, it might be useful for other FreeCAD projects.

# Main features

* **Supported file types:** 
	* .csv - Comma-separated values,
	* .html - HyperText Markup Language,
	* .json - JavaScript Object Notation,
	* .md - MarkDown.

* **Additional features:**
	* export selected spreadsheet or all spreadsheets,
	* custom CSV separator,
	* custom empty cell content,
	* custom CSS decoration for each cell,
	* Qt Graphical User Interface (GUI).

# Screenshots examples

* Qt Graphical User Interface (GUI) main screen with `default settings`:

	![001](https://raw.githubusercontent.com/dprojects/sheet2export/master/Screenshots/001.png)
	
* Info screen at the end about exported files:

	![002](https://raw.githubusercontent.com/dprojects/sheet2export/master/Screenshots/002.png)
	
* Example `html` file output with `3d effect` border decoration:

	![003](https://raw.githubusercontent.com/dprojects/sheet2export/master/Screenshots/003.png)

* The `json` file exemple viewed at [json2table.com](http://json2table.com):

	![004](https://raw.githubusercontent.com/dprojects/sheet2export/master/Screenshots/004.png)

# Known issues

* **Issue**: Too slow for big data like 200 columns x 800 rows.
	* **Workaround**: It takes about 5 minutes on my slow laptop. For the analysis of scientific big data use rather something based on regular expressions, direct XML parsing and mapping. Maybe in the future there will be implemented speed boost feature to improve that but currently I do not have in mind any specific science file type to adjust regular expressions well.

# Contact

Please add all comments and questions to the dedicated
[FreeCAD forum thread](https://forum.freecadweb.org/viewtopic.php?f=22&t=64985).

# License

MIT
