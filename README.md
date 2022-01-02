# Description

This is FreeCAD macro to export spreadsheet. It has been created to support my woodworking project [getDimensions](https://github.com/dprojects/getDimensions).

![001](https://raw.githubusercontent.com/dprojects/sheet2export/master/Screenshots/001.png)

![002](https://raw.githubusercontent.com/dprojects/sheet2export/master/Screenshots/002.png)

![003](https://raw.githubusercontent.com/dprojects/sheet2export/master/Screenshots/003.png)

# Main features

* **Supported file types:** 
	* .csv - Comma-separated values,
	* .html - HyperText Markup Language,
	* .json - JavaScript Object Notation,
	* .md - MarkDown.

* **Additional features:**
	* export selected spreadsheet or all spreadsheets,
	* custom empty cell content,
	* custom CSV separator,
	* CSS table border decoration.

# Known issues

* **Issue**: Too slow for big data like 200 columns x 800 rows.
	* **Workaround**: It takes about 5 minutes on my slow laptop. For the analysis of scientific big data use rather something based on regular expressions, direct XML parsing and mapping. Maybe in the future there will be implemented speed boost feature to improve that but currently I do not have in mind any specific science file type to adjust regular expressions well.

# License

MIT
