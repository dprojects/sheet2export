# Description

<img align="left" src="https://raw.githubusercontent.com/dprojects/Woodworking/master/Icons/sheet2export.png"> This is FreeCAD macro to export spreadsheet to file. It has been created to support my woodworking project [getDimensions](https://github.com/dprojects/getDimensions). However, it might be useful for other FreeCAD projects. 

**Note:** This tool is also part of [Woodworking workbench](https://github.com/dprojects/Woodworking).

<br><br><br><br>

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

# License

[MIT](https://github.com/dprojects/Woodworking/blob/master/LICENSE) for all Woodworking workbench content, so it is more free than FreeCAD.

# Contact

For questions, feature requests, please open issue at: [github.com/dprojects/Woodworking/issues](https://github.com/dprojects/Woodworking/issues)
