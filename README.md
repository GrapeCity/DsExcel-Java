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
## 2.2.0
### New Features
* Export Excel files with shapes to pdf.
* Cut/copy ranges between different workbooks.
* Copy/move worksheet between different workbooks.
* Control whether to adjust page breaks after inserting/deleting rows/columns.
* Customize the row/column/cell delemiter when loading or saving csv file.
* Set the tail repeated rows and right repeated columns when saving to pdf.
* Support for paste options during copy and paste ranges.
* Support IRange.Find() and IRange.Replace() methods.
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
* Support IRange.AutoFit method to fit rows/columns.
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
* Demo Site: [http://demos.componentone.com/gcdocs/gcexceljava/](http://demos.componentone.com/gcdocs/gcexceljava/)
* Maven Repo Address: [https://search.maven.org/artifact/com.grapecity.documents/gcexcel/2.1.0/jar/](https://search.maven.org/artifact/com.grapecity.documents/gcexcel/2.1.0/jar/)
