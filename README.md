# ReadMxd
Open an ArcGIS map document, read properties and output to a text file for easy comparison.

Opening a document in ArcMap and clicking though the layers to find which properties have been set is really tedious, and there is no easy way to compare two documents. This project can help by reading some of the properties and outputting them to a text file. Then two files can be compared in your diff program of choice (KDiff3, Araxis Merge, Compare++, etc.).

ReadMxd can open .mxd, .msd, .aprx, or .lyr files. For .mxd and .lyr, the properties for the document are read one-by-one and some are written out to a log file (the ones that have been added already. Any other props I haven't added yet - feel free to fork this project and put them in!). As .msd and .aprx files are XML-based, each node is read and written to the log.

You can also submit a .txt file that contains a list of any of the above file types for batch processing.

There is also an option to create a spreadsheet with one column for each map document. This is mainly for batch processing of a whole folder full of mxds - the default sheet name is C:\temp\MxdSummary.xlsx, and if found, one more column is added for each mxd read. When finished, you can copy the spreadsheet or just the columns you need to another location and rename it.

Commandline options (all default false):
* -a  All layers. By default, any layers that have visibility false are not shown. Use this option to show properties for them regardless.
* -l  Local log. Put the output props.log file in the same folder as the mxd. Default is <exe path>\MxdProps.log.
* -x  Excel. Add properties to spreadsheet.
* -e  Show full expressions. Long label expressions are truncated to the first 250 chars unless this option is used.
* -s  Read symbols. Show all symbol properties, including all renderers, colours, etc. This can make output file very long.
* -b  Read labels. Show all label and annotation properties.
