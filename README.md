# ReformatRTFMacro

This is a macro I've written using VBA for Microsoft Word to resize and reformat the table cells of a translation project exported as an RTF in six-column-format.\
The code in question is based on a default export of a project created in the CAT tool Déjà Vu, with the Filename column also enabled to be shown on export (i.e. six columns: ID, Source, Target, Comments, Status, and Filename).\
However, it can also be easily adapted for exports that don't include the filename column, as well as for tabular RTF exports from other CAT tools (e.g. MemoQ), provided that the precise requirements (i.e. the number and order of the data columns) are taken into account.\
Examples of the applied macro can be seen in the "before" (Example RTF File for ReformatRTFMacro.docx) and "after" (Example RTF File for ReformatRTFMacro (with Macro Applied).docx) documents uploaded to this repository.\
The code itself can be found in the uploaded .bas file.

## (One Way) To Apply The Macro in Word

1. Enable the Developer tab in Word (File > Options > Customize Ribbon > Tick "Developer" box in right-hand pane).
2. In the ribbon, select Developer > Macros, then add a macro name (e.g. ReformatRTF) in the window that opens, then click "Create".
3. In the code window that opens, paste all the applicable lines of code in between the line with the final apostrophe and the "End Sub" line for the newly created macro (i.e. lines 7-60 in the respository's .bas file).
4. Save the new macro, and it can then be run by going to Developer > Macros again and clicking "Run".

## To Modify The Macro

Lines of code can be removed/commented out or modified as required to make necessary or preferred adjustments to the formatting. For more details on understanding the code, see [Microsoft's VBA Reference guide](https://learn.microsoft.com/en-us/office/vba/api/overview/word).\
\
To provide an overview:\
Each of the selectable columns is assigned a shorter alias for reference (e.g. ActiveDocument.Tables(1).Columns(2) corresponds to the second column from the left in the table in the document, and has been given the name "SourceColumn2" since it corresponds to the source column in this example export).\
ActiveDocument.Tables(1).Rows(1) corresponds to the first row of the table and has been given the alias "HeaderRow" in this case.\
The "With Selection.PageSetup" and subsequent five lines culminating in "End With" reduce the margins to 1cm on all sides of the page.\
The "<ALIAS>.SetWidth ColumnWidth:..." line and following line in each case adjusts the width of the respective column. Changing the number in brackets after "CentimetersToPoints" will change the width that the column is modified to (e.g. CentimetersToPoints(5) = column width of 5cm).\
"<ALIAS>.Select" will then select the respective column/row in Word, allowing global changes to be made to the contents of the cells within those columns/rows.\
"Selection.Font.Size =" changes the size of the font in all cells in the selected column/row (e.g. Selection.Font.Size = 8 = font size 8).\
"Selection.Font.Bold = True" makes the font in all cells in the selected column/row bold. In the present code, this is used to further distinguish the headings in the top row.\
"Selection.Shading.BackgroundPatternColor =" changes the background colour of all cells in the selected column to a colour corresponding to the number following the equals sign. In the present code, "-603917569" corresponds to the first shade of grey (White, Background 1, Darker 5%) selectable from the default palette in Word, and has been applied to the outer two columns. The second shade of default grey "-603923969" (White, Background 1, Darker 15%) is then applied to the header row to further distinguish this from the rest of the table.
