# GrapeCity Documents for Excel, Java Edition
GrapeCity introduces Documents for Excel (GcExcel) Java Edition, a high-speed, feature-rich Excel document API based on VSTO that can help developers work with spreadsheets in Java applications. The library helps to generate, convert to pdf, calculate, format, and parse spreadsheets in any application. You can work with a variety of features like importing spreadsheets, calculate data, query, generate, and export any spreadsheet, add sorting, filtering, formatting, conditional formatting and data validation, grouping, sparklines, charts, shapes, pictures, slicers, comments, hyperlinks, themes etc.  In addition, you can import existing Excel templates, add data and save the spreadsheets back. You can also use GcExcel together with Spread.Sheets, another GrapeCity Spread product that is included in GrapeCity SpreadJS. GcExcel can also import and export Excel template files on the server side. Spread.Sheets works in the browser (client side) as a viewer or editor.

With GcExcel, you can also load, edit, analyse, convert and save spreadsheets in Java applications with full support on Windows, MAC and Linux.

This repository contains source project of Examples and Showcases of GcExcel to help you learn and write your own applications. 

| Directory    | Description    |
| ------------- |-------------|
| gcexcel     | Contains the latest GcExcel jar package and its dependency packages |
| Examples.Library     | A collection of Java examples that help you learn and explore the API features |
| SpringBootDemo/SpringBoot+React     | A source project that demonstrates how to use GcExcel Java with SpringBoot + React + Spread.Sheets|
| SpringBootDemo/SpringBoot+Angular2     | A source project that demonstrates how to use GcExcel with SpringBoot + Angular2 + Spread.Sheets|

# Release Notes
## 3.0.3
### Bugs Fixed
* Exception is thrown when importing xlsx file with certain conditional formatting.(DOCXLS-1844)
* Hyperlink included in ssjson, does not display when exported to PDF. (DOCXLS-1858)
* The column formula is not expanded after adding new rows in Table.(DOCXLS-1865)
* Exception is thrown when importing ssjson with data validations.(DOCXLS-1860)
* When the tag of the copied sheet is changed, the tag of the copy source is also changed.(DOCXLS-1862)
* Exception is thrown when importing ssjson with grouping.(DOCXLS-1882)
* Extra Border is drawn around an image in the header, when xlsx is printed.
* Text styles are lost after copying worksheet and exporting to PDF.(DOCXLS-1814)
* Formulas do not work in generated xlsx file from JSON.(DOCXLS-1847)
* Unable to import specific xlsx file in GcExcel.(DOCXLS-1848)
* Placement.MoveAndSize does not work after exporting to PDF.(DOCXLS-1856)
* External link can not work in the exported PDF when the external file's name contains Chinese characters.(DOCXLS-1867)
* Borders are lost in SpreadJS after importing ssjson generated from GcExcel.(DOCXLS-1885)
## 3.0.2
### Bugs Fixed
* Exception is thrown when importing ssjson that contains duplicated table names.(DOCXLS-1733)
* Null reference exception is thrown when exporting to ssjson.(DOCXLS-1734)
* Index out of range exception is thrown when exporting to pdf.(DOCXLS-1784)
## 3.0.1
### Bugs Fixed
* The font color is changed in SpreadJS after importing the ssjson exported by GCExcel.(DOCXLS-1729)
* The hidden rows and columns are visible in SpreadJS after importing the ssjson exported by GCExcel.(DOCXLS-1729)
* Shape name is changed after copying worksheet.(DOCXLS-1728)
* GcExcel can not load some kind of ssjson correctly. (DOCXLS-1726)
* Some rich text is overlapped after exporting to Pdf.(DOCXLS-1724)
* Exception is thrown during exporting to pdf if the machine is German locale.(DOCXLS-1707) 
* Import ssjson containing data validaiton style and then export,the validation style will be lost in the exported ssjson.(DOCXLS-1692)
* Chart type is changed after importing xlsx file.(DOCXLS-1686)
* Null reference exception is thrown after loading xlsx then exporting to Pdf.(DOCXLS-1684)
* Some shapes are lost after importing ssjson.(DOCXLS-1656)
* "_xlfn" prefix is added before Days formula after exporting to ssjson.(DOCXLS-1655)
* Formula of conditional format is changed after copying a range.(DOCXLS-1636)
## 3.0.0
### New Features
* Templates support to generate Excel reports.
* Support conversion of Excel spreadsheets having slicers to PDF.
* Support new Excel 2016 chart types.
* Support security options when saving to PDF.
* Support document properties when saving to PDF.
* Protect workbook.
* Support chart sheet.
* Support shape with hyperlink.
* Group/Ungroup shapes.
* Calculate outline subtotal.
* Get precedents and dependents of formula cell.
* Add pivot table's grand totals and report layout options similar to Excel.
* Support shape adjustment.
* Support sheet background image to PDF.
* Export Excel files with multiple images to PDF with reduced file size.
* Support license workbook instance.
* Rename pivot fields and Data Fields.
* Support cell tags of GrapeCity SpreadJS.
* Support cell types of GrapeCity SpreadJS.
* Support best fit rows/columns feature of GrapeCity SpreadJS.
### Bugs Fixed
* Unable to set Icon  for IconCriteria.(DOCXLS-1531s)
* "_xlfn" prefix added before IFNA formula while converting to JSON.(DOCXLS-1537)
* The precision of calculated result is incorrect.(DOCXLS-1547)
* Exception is thrown when setting cell value with localDateTime.now().(DOCXLS-7532)
* Text is displayed incompletely in exported PDF document.(DOCXLS-1539)
* DiagonalDown border is lost in exported json.(DOCXLS-1577)
## 2.2.7
### Enhancements
* Customer can detect if there is custom Icon for a cell in specified range.(DOCXLS-1555)
### Bug Fixed
* Null exception is thrown when opening Excel file.(DOCXLS-1377)
* Unable to set Icon  for IconCriteria.(DOCXLS-1531)
* "_xlfn" prefix is added before IFNA formula for ToJSON method.(DOCXLS-1537)
* Characters display incompletely in the exported PDF.(DOCXLS-1539)
* Diagonal down border is lost in the exported ssjson.(DOCXLS-1577)
## 2.2.6
### Enhancements
* Improve the opening time of large spreasheet when it contains shape and auto-fit rows.
* JP built-in named styles are supported.
### Bug Fixed
* The formula result is not precise which leads to the conditional checking failed.(DOCXLS-1453)
* The location and size of shapes are incorrect after opening the saved Excel under JP culture.(DOCXLS-1361)
* The content of exported pdf is incorrect.(DOCXLS-1315)
## 2.2.5
### Bug Fixed
* The picture name is changed after saving a workbook to json several times.(DOCXLS-1427)
* Corrupted Excel file is generated after copying a range containing checkbox control.(DOCXLS-1446)
* Cell border is lost in exported pdf file..(DOCXLS-1450)
* The auto row height is incorrect in the exported pdf if there is a merged cell and wrap text is true.(DOCXLS-1429)
* Exception is thrown when opening a Excel file with a table whose name contains '.'.(DOCXLS-1451)
* The marker of data series is changed after saving and loading json.(DOCXLS-1444, DOCXLS-1273)
* The  content should not be wrapped cell's value is a number.(DOCXLS-1430)
## 2.2.4
### Bug Fixed
* IndexOutOfRange exception is thrown when the same XLSX file is opened with parallel threads.(DOCXLS-1368)
* Exception is thrown after loading ssjson and then save to PDF file.(DOCXLS-1382)
* Column width changed after loading ssjson.(DOCXLS-1371)
* Unable to fill shape with picture through API.(DOCXLS-1376)
* Exception is thrown after loading ssjson and then invoke SetColumnWidth.(DOCXLS-1378)
* Some text can not be rendered completedly in the exported PDF file.(DOCXLS-1404)
* Unable to set a picture fill for shape.(DOCXLS-1376)
* Exception is thrown when saving a worksheet to PDF file two time.(DOCXLS-1415)
* The calculation result of INDIRECT function is different with Excel.(DOCXLS-1335)
## 2.2.3
### Bug Fixed
* Incorrect sparkine is rendered in the exported PDF file.(DOCXLS-1242)
* InvalidFormulaException is thrown during loading ssjson.(DOCXLS-1266)
* An error occurs while saving to PDF file.(DOCXLS-1303)
* Out of memerory occurs while loading an Excel file.(DOCXLS-1279)
* Stack overflow occurs while saving to PDF file.(DOCXLS-1315)
* Exception is thrown while loading a xlsm file.(DOCXLS-1335)
* Calculation results from GCExcel doesn't match the results from MSExcel.(DOCXLS-1322, DOCXLS-1323)
* Cell Style is not correct after GcExcel toJSON and then SpreadJS fromJSON.(DOCXLS-1334)
## 2.2.2
### Enhancements
* Customer can overwrite current function implementation when registering a new custom function.(DOCXLS-1146)
### Bug Fixed
* Exception is thrown when exporting to json.(DOCXLS-1120)
* Exception is thrown after adding a lot of hyperlinks then saving to Excel.(DOCXLS-1154)
* Some text can not display completely in the exported pdf file.(DOCXLS-1156)
* ArgumentOutOfRangeException is thrown on using Calculate method for a Workbook.(DOCXLS-1211)
* If calculated result is too large(20+ digits) in the cell, the value acquired by GcExcel is horizontally reversed.(DOCXLS-1127)
* There is a difference between the Excel display value and the Text value acquired by GcExcel.(DOCXLS-1128)
* When a cell referenced by a formula is moved by "Cut" method, the formula is not automatically updated.(DOCXLS-1210)
* The exported pdf and Excel file is corrupted when using GcExcel in multiple threading. (DOCXLS-1208)
## 2.2.1
### Bug fixed
* GcExcel throws exception when loading ssjon with invalid image source.(DOCXLS-1115)
* Exception is thrown when opening an Excel with a lot of conditional formatting.(DOCXLS-1047)
* Unable to export to PDF with a specified culture.(DOCXLS-1043)
* The exported Excel file does not open correctly after loading ssjson.(DOCXLS-1076)
* Redundant lines are rendered while exporting to PDF.(DOCXLS-1068)
* SSJSON is not exported correctly if there are charts with duplicate names.(DOCXLS-1067)
* PrintInfo.firstPageNumber in ssjson is always 1 even though it is not set for IWorksheet.PageSetup.(DOCXLS-1107)
* Exception is thrown when opening certain Excel file.(DOCXLS-1104)
* Exception is thrown when loading a ssjson containing font with line height.(DOCXLS-1110)
* Leading apostrophe in a cell value is displayed as part of the cell text (not consistent with Excel behavior).(DOCXLS-1041)
* Exception is thrown when loading ssjson with custom function.(DOCXLS-1044)
* Copy behavior is inconsistent with Excel when the target range is larger than the source range.(DOCXLS-844)
* Exception is thrown after opening certain Excel file and then export to ssjson.(DOCXLS-1116)
* GcExcel can not export text with Simhei bold font.(DOCXLS-937).
* Load ssjson with combobox cells and then export to Excel, the cell value of combobox cell is wrong in the exported Excel file.(DOCXLS-1106)
## 2.2.0
### New Features
* Export Excel files with shapes to pdf.
* Cut/copy ranges between different workbooks.
* Copy/move worksheet between different workbooks.
* Control whether to adjust page breaks after inserting/deleting rows/columns.
* Customize the row/column/cell delemiter when loading or saving csv file.
* Set the tail repeated rows and right repeated columns when saving to pdf.
* Support for paste options during copy and paste ranges.
* Support IRange.find() and IRange.replace() methods.
* Show or hide different kinds of pivot table styles.
* Support for exporting pivot table styles to pdf.
* Support setting the number format for each pivot field.
* Japanese ruby can now be preserved after Excel I/O.
* Get and customize each page settings before saving to pdf file.
* Render any sheet range inside a pdf file.
* Keep rows/columns togother when saving to pdf.
* Save multiple workbooks to one pdf file.
* Export specific pages from spreadsheet to pdf
* Save multiple spreadsheet pages into one pdf page.
* Support IRange.autoFit method to fit rows/columns.
* Support IRange.FormulaArrayR1C1 property to get or set array formula in R1C1 format.
* Support more import flags during open an Excel file.
* OLEObject will be preserved after Excel I/O.
* Support Shrink to fit, wrapped text while saving to pdf.
### Bug fixed
* GcExcel ignores 'ignore_empty' parameter in TEXTJOIN formula.(DOCXLS-970)
* Large JSON file generated when using ToJson on a particular Workbook.(DOCXLS-968)
* UsedRange.Value sets improper values to the range when Formula is set to Empty.(DOCXLS-956)
## 2.1.5
### Bug Fixed
* The column width, font size and picture's left position changes after JSON I/O with GcExcel.(DOCXLS-902)
* The cell font changes after JSON I/O with GcExcel.(DOCXLS-921)
* Visible property of secondary category axis returns incorrect value.(DOCXLS-849)
* The major and minor units of value axis of stacked area chart return incorrect value.(DOCXLS-805)
* The major and minor units of value axis for 100% stacked chart return incorrect value.(DOCXLS-800)
* The data label's text returns incorrect value when Axis.DisplayUnit is set. (DOCXLS-768)
* Exception is thrown on opening json file that contains invalid formula.(DOCXLS-948)
## 2.1.4
### Enhancements
* Performance of exporting a single worksheet to PDF was improved significantly.
* fromJson method  was improved when json contained table style and multiple named styles.
* GcExcel now sets minimum value of Zoom factor to 10% when exporting spreadsheet to PDF, similar to MS Excel setting. 
### Bug Fixed
* The TintAndShade property does not work as expected if Color is set.(DOCXLS-872)
* Cells containing SUBTOTAL formula do not return correct values.(DOCXLS-881)
* InvalidFormulaException is thrown when opening Excel files saved by Open XML SDK.
## 2.1.3
### Bug fixed
* Image size specified in code does not apply on the image in the generated excel file (DOCXLS-787)
* Setting formulas using @ symbol trims the column name.(DOCXLS-804)
*  Data label's separator does not return correct value.(DOCXLS-830)
*  Cell's value always return 0 when workbook.getEnableCalculation() is false.(DOCXLS-834)
*  Improve the performance of lookup functions when looking up string values.
## 2.1.2
### Bug fixed
* GcExcel is unable to load an Excel file consisting SUMPRODUCT formula.(DOCXLS-733)
## 2.1.1
### Enhancements
* Added IWorksheet.FixedPageBreaks to control whether to adjust page breaks after inserting/deleting rows or columns.
* Improved the effect of wrapped text to pdf.
### Bug fixed
* Page breaks are not adjusted after inserting/deleting rows or columns.(DOCXLS-728)
* NullReferenceException is thrown after calling IRange.ClearContents().(DOCXLS-731)
* NullReferenceException occurs on editing and saving XML generated by ClosedXML.(DOCXLS-726)
* Refreshing a PivotTable through code and saving it, corrupts the generated Excel file.(DOCXLS-739)
* ArgumentOutOfRangeException thrown on refreshing PivotTable present in some *.xlsx file.(DOCXLS-740)
* On deleting the active sheet and saving the Excel file, an error dialog is observed on opening the saved Excel file.(DOCXLS-747)
* StackOverflowException is thrown on invoking SetSourceData for Chart.(DOCXLS-724)
* Color.Empty leaves black color on a cell.(DOCXLS-722)
* Error is thrown while using FromJson method.(DOCXLS-734)
* Using MAXIFS function results in #NAME? error in cell.(DOCXLS-749)

## 2.1.0
### New Features
* The performance of workbook.fromJson() method has been enhanced when the JSON file contains multiple styles.
* Users can now import and export spreadsheets that contain macros. While these will not be executed, the macros will now be preserved when saving.
* The support for loading and saving GrapeCity SpreadJS JSON files with shapes have been added.
* Users can now set rich text format in the cells by applying different styles to the textual information entered in the cell.
* While working with custom named styles, users can now modify an existing style and add it to the Styles collection.
* Users can now export Excel files with vertical text to PDF.
* Now, users can insert any background image to the worksheet including their organization logo, custom watermark or a wallpaper of their choice without any issues.
* PDFBox can now be installed automatically for all versions of Eclipse Maven plugin.
* Extensive support for the new Date Time API that has been introduced recently with JDK 8 has been provided.
* The pivot table has been enhanced in order to support the date field group in Excel 2016.
* Some overloads have been added for Open and Save methods to avoid passing file format.
### Bug fixed
* The Workbook.calculate() method now evaluates the cell values correctly.
* While saving an Excel file to open XML format, the logical value of the cell is now calculated without any errors.

## 2.0.1
### Enhancements
* Improved the performance of Workbook.fromJson method, when json file contains lot of styles.
* PDFBox now can be installed automatically for all versions of eclipse maven plugin.
### Bug fixed
* Exception thrown when Workbook opens the stream returned by HttpServletRequest.getInputStream().
* Null pointer exception is thrown on saving to PDF, if the font used is null.
* GcExcel throws exception on loading ssjson file with null values.
* Merged range in table couldn't be rendered to pdf.
* The hidden rows are still rendered to pdf after loading ssjson file.


# Other Resources
* Product Home Site: [https://www.grapecity.com/documents-api-excel](https://www.grapecity.com/documents-api-excel)
* Demo Site: [https://www.grapecity.com/documents-api-excel/demos/](https://www.grapecity.com/documents-api-excel/demos/)
* Maven Repo Address: [https://search.maven.org/artifact/com.grapecity.documents/gcexcel/](https://search.maven.org/artifact/com.grapecity.documents/gcexcel/)
