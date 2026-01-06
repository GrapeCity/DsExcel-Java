# Document Solutions for Excel, Java Edition
Programmatically generate, load, modify, and convert Excel .xlsx spreadsheets with ease in your Java applications. Document Solutions for Excel (DsExcel) is a cross-platform high-speed, small-footprint spreadsheet API library that requires no dependencies on Excel. Applications using this spreadsheet API can be deployed to the cloud, Windows, Mac, or Linux. The powerful calculation engine and breadth of features means you'll never have to compromise on design or requirements.

At a glance:
- Create, load, edit, and save Excel spreadsheets
- Save to .XLSX, PDF, Image, HTML, CSV, and JSON
- Based on the Excel Object Model with zero Excel dependencies
- Deploy locally, inhouse or to Cloud including Azure and AWS
- 450+ Excel Functions and custom functions supported to perform complex calculations
- Use the Templates to create custom Excel reports
- 2x+ faster and less memory than Apache POI

Based on the extensive Excel Object Model, the interface-based API allows you to import, calculate, query, generate, and export any spreadsheet scenario. With the VS Tools for Office-style API, you can create custom styles using the same elements as VS Tools for Office.

Using DsExcel, you can create full reports, sorted and filtered tables, sorted and filtered pivot tables, dashboard reports, add charts, slicers, sparklines, conditional formats, import and export Excel templates, convert spreadsheets to PDF and so much more.

DsExcel comes with a full-featured [JavaScript Data viewer control (DsDataViewer)](https://www.npmjs.com/package/@mescius/dsdataviewer).

If you'd like to remove the trial watermark and other [trial limitations](https://developer.mescius.com/document-solutions/dot-net-excel-api/docs/online/LicenseInformation.html), please email us.sales@mescius.com to request your 30-day evaluation key.

**Complete Client-Server Spreadsheet Solution**

You can optionally integrate DsExcel with the SpreadJS JavaScript spreadsheet as a client-side editor/viewer solution when working with Excel files for a complete client-server solution. View the complete supported features list [here](https://developer.mescius.com/documents-api-excel-java/docs/online/support-for-spreadjs-features.html) or download a trial from [NPM](https://www.npmjs.com/package/&#64;grapecity/spread-sheets) or the [SpreadJS](https://developer.mescius.com/spreadjs) page.

## Resources

- [Download 30-day Evaluation](https://developer.mescius.com/documents-api-excel-java/download)
- [Product Page](https://developer.mescius.com/document-solutions/java-excel-api)
- [Online Demo](https://developer.mescius.com/document-solutions/java-excel-api/demos/)
- [DsDataViewer Demo](https://developer.mescius.com/document-solutions/javascript-data-viewer/demos/Overview)
- [Getting Started](https://developer.mescius.com/document-solutions/java-excel-api/docs/online/getting-started.html)
- [License Information](https://developer.mescius.com/document-solutions/java-excel-api/docs/online/LicenseInformation.html)
- [Licensing FAQ](https://developer.mescius.com/document-solutions/licensing)
- [How to get Trial Keys](https://developer.mescius.com/document-solutions/java-excel-api/docs/online/LicenseInformation.html)
- [DsExcel Blogs](https://developer.mescius.com/blogs/categories/products/documents)
- [Online Documentation](https://developer.mescius.com/document-solutions/java-excel-api/docs/online/overview.html)
- [Offline Documentation (PDF)](https://developer.mescius.com/document-solutions/java-excel-api/docs/offlinehelp.pdf)

## Other Document Solutions API Solutions

- [Document Solutions for Excel, .NET](https://developer.mescius.com/document-solutions/dot-net-excel-api)
- [Document Solutions for PDF](https://developer.mescius.com/document-solutions/dot-net-pdf-api)
- [Document Solutions PDF Viewer](https://developer.mescius.com/document-solutions/javascript-pdf-viewer)
- [Document Solutions for Word](https://developer.mescius.com/document-solutions/dot-net-word-api)
- [Document Solutions for Imaging](https://developer.mescius.com/document-solutions/dot-net-imaging-api)

This repository contains source project of Examples and Showcases of DsExcel to help you learn and write your own applications. 

| Directory    | Description    |
| ------------- |-------------|
| dsexcel     | Contains the latest DsExcel jar package and its dependency packages |
| Examples.Library     | A collection of Java examples that help you learn and explore the API features |
| SpringBootDemo/SpringBoot+React     | A source project that demonstrates how to use DsExcel Java with SpringBoot + React + Spread.Sheets|
| SpringBootDemo/SpringBoot+Angular2     | A source project that demonstrates how to use DsExcel with SpringBoot + Angular2 + Spread.Sheets|

---
## 9.0.0
## Added
* Support Lossless of Query Table.(DOCXLS-11502)
* Added support for new Excel functions VALUETOTEXT and ARRAYTOTEXT.(DOCXLS-12696)
* Support controlling SharedFormula export in XLSX export.(DOCXLS-12881)
* Support AI functions.(DOCXLS-13155)
* Support SJS/SSJSON IO of Threaded comments.(DOCXLS-13331)
* Performance optimization.(DOCXLS-13478)

## Fixed
* Copying a Worksheet with many formulas is very slow.(DOCXLS-4633)
* Opening and profiling a Workbook is noticeably slower.(DOCXLS-11077)
* Saving a Workbook to PDF can trigger a console error.(DOCXLS-12187)
* Formatting and AutoFit on very large reports causes severe slowdowns and high memory use.(DOCXLS-12607)
* Copying formulas across Worksheets is significantly slower in newer versions.(DOCXLS-12672)
* Range.Copy on very large ranges can be extremely slow.(DOCXLS-12709)
* Enabling the calculation engine after sheet and named-range changes can throw a NullReferenceException.(DOCXLS-13383)
* Opening certain Workbooks can fail on JDK 21+.(DOCXLS-13397)
* XLOOKUP can return incorrect numeric results compared to Excel.(DOCXLS-13433)
* Calling Range.getText() can throw an exception.(DOCXLS-13443)
* LET formulas adding a scalar to very large arrays return values only in the first row and zeros elsewhere.(DOCXLS-13449)

## 8.2.5
## Fixed
* Calling waitForCalculationToFinish can block for a long time during Workbook calculation.(DOCXLS-13207)
* Workbook.Open can take an excessively long time to complete.(DOCXLS-13264)
* Enabling CanShrinkToFitWrappedText when exporting to PDF can cause the save operation to fail.(DOCXLS-13265)
* Formulas that combine LET, TAKE, BYROW, and LAMBDA can evaluate to #VALUE!.(DOCXLS-13270)
* Formula evaluation under Japanese culture can produce incorrect values compared to Excel.(DOCXLS-13279)
* A PivotTable created from a Data Model with no Rows or Columns selected can disappear or result in a corrupted XLSX after Excel I/O.(DOCXLS-13330)
* Diagonal cells can be lost after processing a Template (ProcessTemplate).(DOCXLS-13341)
* Range values can become #N/A errors after copying a Range between Workbooks.(DOCXLS-13380)
* Image fill applied to a Comment shape can be ignored while the same image displays correctly on a Worksheet shape.(DOCXLS-13382)
## 8.2.4
## Fixed
* Destination Cells lose Data Validation when rows are copied.(DOCXLS-12951)
* Embedded images in Picture-In-Cell disappear after Excel I/O.(DOCXLS-13165)
* Exported Workbooks containing a Chart with a Secondary Axis cause an error when opened in Excel.(DOCXLS-13179)
* HasArray and CurrentArray do not recognize Dynamic Array Formulas.(DOCXLS-13183)
* Only the first Cell expanded from a JSONTable participates in calculation.(DOCXLS-13184)
* Certain saved XLSM files are reported as corrupted by Excel.(DOCXLS-13192)
* Formulas produced during Template expansion can be incorrect.(DOCXLS-13202)
* IMSUM and IMPRODUCT return zero when referencing Cells that contain complex numbers.(DOCXLS-13205)
* XLOOKUP can return #N/A when an Error-type if_not_found argument is supplied even though the lookup value exists.(DOCXLS-13206)
* Exception is thrown on Template processing when the ds Data Source is a plain JSON object.(DOCXLS-13212)
* Sorting a Range copied between Workbooks can fail with an exception.(DOCXLS-13215)
* IFS returns incorrect results when its first argument is provided as an Array.(DOCXLS-13216)
* Comments attached to expanded Cells are lost during Template expansion.(DOCXLS-13220)
## 8.2.3
## Fixed
* Conditional formatting in the sjs extLst node is not parsed, causing differences in cell text rendering compared to SpreadJS.(DOCXLS-13107)
* Inserted shapes may be misaligned due to oversized anchor offsets exceeding column/row boundaries.(DOCXLS-13112)
* Enabling WrapText increases row height unexpectedly from the second row when customHeight is true.(DOCXLS-13114)
* Formula evaluation may produce incorrect numeric results under a German locale.(DOCXLS-13131)
* Exporting a Range with quoted conditional formatting values causes an InvalidFormulaException during Image or PDF export.(DOCXLS-13133)
* Shape hyperlinks are lost when exporting and reloading a Workbook as JSON.(DOCXLS-13152)
## 8.2.2
## Fixed
* File Upload cell values are lost during copy/paste or export, causing missing file names.(DOCXLS-11384)
* Setting a data source can overwrite cell content and remove formulas.(DOCXLS-12468)
* First update of BackgroundPicture width may be ignored after changing FontSize.(DOCXLS-12614)
* Freeze Pane position shifts when file contains collapsed groups.(DOCXLS-12784)
* Workbook.Calculate() may return errors if the calculation engine starts Off.(DOCXLS-13053)
* Full-width "＆" in Worksheet names change to half-width "&" in formulas after save/load.(DOCXLS-13073)
* Upgrading can cause serious performance regression in Workbook calculation.(DOCXLS-13074)
* Second Workbook calculation may produce incorrect results.(DOCXLS-13075)
* PivotTable Slicers export incorrectly to PDF.(DOCXLS-13077)
* Certain Excel templates may throw "Stack empty" exception.(DOCXLS-13080)
* Inserting a row at the top can move Header Filter icons into the data row.(DOCXLS-13085)
* Resaving Dynamic Array formulas may fill spill range with values and cause #SPILL!.(DOCXLS-13110,DOCXLS-13111)
## 8.2.1
## Fixed
* Chart axis label spacing is ignored when exporting to image, showing dense labels instead.(DOCXLS-11876)
* Localized color tokens(e.g., [Red] in Chinese) cause number format errors or inconsistencies.(DOCXLS-12795)
* Extra blank comment may appear in exported XLSX files(visible in WPS).(DOCXLS-12847)
* Percentage format is lost after cell expansion.(DOCXLS-12875)
* Setting a cell value can throw System.ArgumentException.(DOCXLS-12894)
* Opening large .ssjson files with many styles is very slow.(DOCXLS-12923)
* Some dynamic array formulas return incorrect results on first calculation, correct only after recalculating.(DOCXLS-12924)
* Exporting certain workbooks to HTML can throw Invalid range errors(e.g., empty sheet with PrintArea).(DOCXLS-12931)
* Chart axis bounds and intervals change when exporting to PDF.(DOCXLS-12942)
* DsExcel calculation results differ from Excel and may need multiple recalculations.(DOCXLS-12944)
* Page headers with "%%%" cause errors and prevent saving/printing.(DOCXLS-12949)
* Dynamic array formulas starting with + and cross-sheet references return #VALUE!.(DOCXLS-12952)
* Pivot cache and binding data remain in JSON after deleting PivotTable and source sheet.(DOCXLS-12953)
* Saving under Japanese locale does not preserve certain date number formats.(DOCXLS-12955)
* Saving SJS files with empty RichText entries in SharedStrings.json throws exceptions.(DOCXLS-12987)
## 8.2.0
## Added
* Added support for new Excel functions GROUPBY, PIVOTBY, and PERCENTOF.(DOCXLS-9486)
* Introduced the Evaluate2 API to return spilled values from dynamic array formulas.(DOCXLS-11502)
* Added the ability to get and set Alternative text for shapes.(DOCXLS-11891)
* Enabled setting chart sheet zoom to "Fit to Selection".(DOCXLS-11930)
* Custom functions can now accept error parameters.(DOCXLS-12010)
* Added support for the new Excel formula TRIMRANGE.(DOCXLS-12136)
* Enabled retrieval of sheet count and names for GanttSheets from SJS files.(DOCXLS-12206)
* Expanded the ShapeType enum with new types.(DOCXLS-12226)
* Improved the performance of template expansion when the data source is ITableDataSource and the ExpansionType is list mode.(DOCXLS-12230)
* Support for Eta-Reduced Lambda.(DOCXLS-12446)

## Fixed
* Exporting to Excel takes unusually long when files contain multiple VLOOKUP functions.(DOCXLS-10630)
* Font formatting changes unexpectedly after opening and saving files.(DOCXLS-11791)
* Memory usage is excessively high when processing large-scale data with Workbook.processTemplate().(DOCXLS-11909)
* Report rendering speed is slower than expected when processing large datasets.(DOCXLS-11950)
* Performance differs significantly between Template and setDataSource methods when rendering large lists.(DOCXLS-12216)
* Processing templates with a large number of rows takes longer than expected and may cause memory issues in multi-threaded scenarios.(DOCXLS-12244)
* Font formatting is inconsistent when converting SJS files to XLSX, resulting in font changes.(DOCXLS-12504)
* Executing the toImage function on Mac and Linux triggers console errors due to unsupported fonts or emoji rendering issues.(DOCXLS-12539)
* Freeze pane location shifts after processing files containing collapsed groups, causing incorrect frozen areas.(DOCXLS-12784)
* Sorting order for Japanese text in Pivot Tables does not follow locale-specific rules.(DOCXLS-12801)
* Saving certain workbooks fails with System.ArgumentOutOfRangeException due to issues in chart handling.(DOCXLS-12810)
* Table filter buttons appear inconsistently compared to older versions, indicating lost filter information in processed templates.(DOCXLS-12860)

## 8.1.5
## Fixed
* When a specific file is loaded and borders are added to cells, the used range becomes incorrect.(DOCXLS-8994, DOCXLS-12767)
* A workbook calculation exception occurs when a formula contains an external workbook reference.(DOCXLS-12129)
* Unexpected file size increases when applying NumberFormat and HorizontalAlignment to entire rows or columns.(DOCXLS-12650)
* The LET function returns a #REF! error.(DOCXLS-12697)
* The chart does not meet expectations after data expansion.(DOCXLS-12703)
* Issue with the HOUR formula.(DOCXLS-12706)
* After setting multiple SubtotalTypeNone options in a pivot table, exporting Excel SubtotalTypeNone settings does not take effect.(DOCXLS-12721)
* An InvalidFormulaException occurs when a LAMBDA parameter is named "rc" in a Name Manager formula.(DOCXLS-12763)
* Lack of parallelism in the Calculation Engine for independent workbook instances.(DOCXLS-12776)
## 8.1.4
## Fixed
* After ignoring formulas, the calculation results of LAMBDA formulas are not exported properly.(DOCXLS-12558)
* A NullPointerException is thrown when saving a Workbook to SJS, HTML, or PDF formats.(DOCXLS-12585)
* After setting FitToPagesWide, the exported PDF file still has two pages.(DOCXLS-12588)
* The calculated result is incorrect when ROW function is used in defined name.(DOCXLS-12626)
## 8.1.3
## Fixed
* Grouping a range results in an incorrect outline.(DOCXLS-12484)
* The result of setting the option to insert entire rows and columns is not as expected.(DOCXLS-12488)
* An exception is thrown when exporting a PDF file that contains a specific chart.(DOCXLS-12491)
* Fetching IconType throws an ArrayIndexOutOfBoundsException when the icon is IconType.NoCellIcon.(DOCXLS-12503)
* Formula cells are not evaluated after calling Workbook.Calculate() method.(DOCXLS-12536)
* Pivot table totals to an incorrect value.(DOCXLS-12545)
## 8.1.2
## Fixed
* The applyFilter method does not take effect.(DOCXLS-12106)
* Exception occurs in Workbook.processTemplate().(DOCXLS-12209)
* The PivotTable data source is not updated after the sheet name is changed.(DOCXLS-12227)
* An error is thrown when saving an SJS file.(DOCXLS-12281)
* The program throws an exception when opening an SSJSON file.(DOCXLS-12286)
* Issues with SUMIF calculation when exporting Excel to PDF.(DOCXLS-12307)
* ConditionalFormats values are lost after processing the template.(DOCXLS-12315)
* The value of RowCount is not expected.(DOCXLS-12316)
* The exported PDF file does not display consistently with SpreadJS.(DOCXLS-12317)
* Prompt for illegal parameters when importing an SSJSON file.(DOCXLS-12320)
* Differences in the image output of chart generated by DsExcel and Microsoft Excel.(DOCXLS-12322)
* Fetching IconType throws NullPointerException.(DOCXLS-12335)
* The result of the IRange.replace method is incorrect.(DOCXLS-12349)
* Shapes are not output correctly when exporting to PDF.(DOCXLS-12430)
* Exception occurs when saving a specific Excel file.(DOCXLS-12440)
## 8.1.1
## Fixed
* The result of a dynamic array formula is incorrect when there is a default value in the worksheet.(DOCXLS-11985, DOCXLS-12145)
* The exported Excel file is corrupted when the workbook contains a particular Pivot Chart.(DOCXLS-12097)
* An exception is thrown when executing the getRefersTo method.(DOCXLS-12157)
* The column positions are incorrect after the template is extended.(DOCXLS-12158)
* The count of the runs of rich text increases after calling the IRange.characters() method.(DOCXLS-12185)
* The result of the IFont.getUnderline() method is incorrect in rich text.(DOCXLS-12186)
* An exception is thrown when exporting an Excel file containing a 3D reference.(DOCXLS-12193)
* The font in the exported SSJSON file is corrupted.(DOCXLS-12202)
* The items in the dropdown list in the exported SJS file are changed.(DOCXLS-12207)
* A #VALUE! error occurs when using a dynamic array formula.(DOCXLS-12210)
* Camera shapes appear blurry in the exported PDF file.(DOCXLS-12218)
* An exception is thrown when calling the processTemplate() method.(DOCXLS-12239)
* An exception occurs when exporting the 101st file, even though it is licensed.(DOCXLS-12272)
* The cell format is lost in the exported Excel file.(DOCXLS-12273)
* The used range in the exported Excel file is incorrect when the workbook contains a dynamic array formula.(DOCXLS-12285)
## 8.1.0
## Added
* Added Clone method to duplicate the workbook instance.(DOCXLS-6685)
* Support for adding cell addresses to exported tables in HTML.(DOCXLS-11173)
* Performance Improvements for VLOOKUP and UNIQUE functions.(DOCXLS-11501)
## Fixed
* Exception is thrown on calling IWorksheet.setDataSource() method.(DOCXLS-4644)
* The sorting result of the pivot table is incorrect in the exported PDF file.(DOCXLS-8864)
* Performance issue on calculating the UNIQUE function.(DOCXLS-10644)
* The custom object is lost in the exported Excel file.(DOCXLS-11387)
* Performance issue on calculating the VLOOKUP function.(DOCXLS-11392)
* The address of hyperlink gets changed after processing the template.(DOCXLS-12150)

## 8.0.5
## Fixed
* Exception is thrown on saving an SJS file after deleting a worksheet which is referenced by a linked picture.(DOCXLS-10752)
* The diagonal border in the exported PDF file is incorrect.(DOCXLS-11792)
* The Y axis is not shown after adding it using API.(DOCXLS-11964)
* The result of CODE function is incorrect when the parameter contains CJK characters.(DOCXLS-11987)
* The result of IWorksheet.getUsedRange() is incorrect when UsedRangeType contains Axis.(DOCXLS-12003)
* The binding path of table is lost in the exported SJS file after processing template.(DOCXLS-12009)
* The active sheet is incorrect in the exported SJS file.(DOCXLS-12033)
* The font in shape is incorrect in the exported Excel file.(DOCXLS-12039)
* Exception is thrown on loading a specific Excel file.(DOCXLS-12064)
* The name reference changed in the exported Excel file when name refers to external workbook.(DOCXLS-12076)
* The result of IRange.getText() is incorrect.(DOCXLS-12081)
* Exception is thrown on loading particular SSJSON file.(DOCXLS-12085)
* Exception is thrown on loading particular Excel file with extenal reference.(DOCXLS-12086)
* The style name is incorrect when using the DsExcel API.(DOCXLS-12093)
* The value is lost when setting IncludeFormula is set to false when opening an SJS file.(DOCXLS-12096)
* There are extra cell border in the exported PDF file when the workbook contains both merge and automerge cells.(DOCXLS-12099)
* Performance issue occurs on calculating workbook contains SORT function.(DOCXLS-12101)
* Exception is thrown on calling Workbook.Calculate() method when the last parameter of NPV function is empty.(DOCXLS-12142)
## 8.0.4
## Fixed
* An exception was thrown when loading a specific SSJSON file.(DOCXLS-11827)
* The result was incorrect after calling the IRange.sort() method.(DOCXLS-11865)
* The binding information for table columns was lost after processing the template.(DOCXLS-11882)
* The formula in the exported workbook was incorrect after processing the template.(DOCXLS-11919)
* The custom formatter was lost in the exported SJS file.(DOCXLS-11941)
* The shape's border remained visible in the exported PDF and SJS files despite being set to invisible.(DOCXLS-11961)
* The custom formatting of the table was lost in the exported Excel file.(DOCXLS-11962)
* The TintAndShade setting did not take effect.(DOCXLS-11970)
## 8.0.3
## Fixed
* The text display different comparing with Excel in the exported PDF file.(DOCXLS-11740)
* Exception is thrown on processing Template.(DOCXLS-11799)
* Unexpected cell borders appear in the exported PDF file.(DOCXLS-11834,DOCXLS-11835)
* The scale of tick labels in vertical axis is incorrect in the exported PDF file.(DOCXLS-11837)
* Exception should be thrown when data source in template is missing.(DOCXLS-11856)
* There are Spill and NA Error when using COUNTIFS formula.(DOCXLS-11859,DOCXLS-11860)
* The cell locked status has changed after exporting workbook to SSJSON file then import it.(DOCXLS-11866)
* The exported Excel file gets corrupted if result of formula contains ampersand.(DOCXLS-11873,DOCXLS-11907)
* The reference in the XLOOKUP function is not expected after processing Template.(DOCXLS-11874)
* StackOverflowError when data containing specific date is used in pivot table.(DOCXLS-11884)
* The result of INDEX function is incorrect.(DOCXLS-11904)
* The expanded result is not as expected after processing Template.(DOCXLS-11914)
* The behavior or autofilter is incorrect.(DOCXLS-11922)
* The form controls in the exported Image/PDF file are incorrect.(DOCXLS-11934)
## 8.0.2
## Fixed
* The cell reference is not as expected in the expanded result after processing template.(DOCXLS-10967)
* The size of Pie chart is smaller than Excel in the exported Image file.(DOCXLS-11272)
* Performance issues on processing templates with a large expanded result range.(DOCXLS-11490)
* The cell tag is lost in the exported Excel file after processing template.(DOCXLS-11590)
* The row count is incorrect in the exported SJS file.(DOCXLS-11728)
* Exception is thrown on copying cell across workbook if the source range contains conditional format.(DOCXLS-11735)
* Exception is thrown on exporting to the PDF file when workbook uses Cantarell fonts.(DOCXLS-11739)
* Hidden data labels of a chart are exported to the image if repeated category names are present(DOCXLS-11742)
* Grouped Pivot Slicers get lost when Excel file is loaded and saved.(DOCXLS-11753)
* Extra underline appears in the exported PDF file.(DOCXLS-11759)
* Performance issue on calculating big Excel file with lots of SUMIF function.(DOCXLS-11770)
* Exception is thrown on calling IFormatCondtions.Count when the conditional format is in a Pivot Table.(DOCXLS-11776)
## 8.0.1
## Fixed
* The cell value is incorrect when it is in the spilled range of dynamic array formula.(DOCXLS-11391)
* The style in the hidden column is shown after processing template.(DOCXLS-11489)
* The z-order of shapes changed after processing template.(DOCXLS-11535)
* Exception is thrown on calling the processing template method when null exists in the data source.(DOCXLS-11588)
* Background picture is not copied when calling IRange.Copy method.(DOCXLS-11612)
* The result of XIRR formula is incorrect.(DOCXLS-11614)
* Exception is thrown on getting value when cell uses PERCENTRANK.EXC formula.(DOCXLS-11624)
* Some properties of chart data label lost after Excel I/O.(DOCXLS-11658)
* The font of zysteel.ttf is not rendered correctly in the exported PDF file.(DOCXLS-11676)
* The direction of arrow shape are not rendered correctly in the PDF file.(DOCXLS-11687)
* The result of COUNTIFS formula is incorrect with "" and "<>" criteria.(DOCXLS-11689)
* Exception is thrown on loading SSJSON file when activeRow and activeCol are null.(DOCXLS-11696)
* Exception is thrown on loading an XLSX file.(DOCXLS-11697)
* The effect of MediumDashed border is incorrect in the exported PDF file.(DOCXLS-11698)
* The legend of chart is incorrect after processing template.(DOCXLS-11699)
* The result of dynamic array formula is incorrect after processing template.(DOCXLS-11700)
* The conditional format is not expanding as expected after processing template.(DOCXLS-11701)
* Authorization error occurs when the device name is in CJK culture.(DOCXLS-11703)
* The number format is incorrect in the exported XLSX file after calling IRange.Copy method.(DOCXLS-11706)
* Some rich text is incorrect after Excel I/O.(DOCXLS-11708)
* Exception is thrown on saving SJS file when the reference of filter is empty.(DOCXLS-11725)
* Exception is thrown on saving JSON file when workbook contains specific custom object.(DOCXLS-11733)
* The content of cell is incorrect if cell width is not enough in the exported PDF file.(DOCXLS-11741)
## 8.0.0 
## Added
* Import data from object collections and Data Tables.(DOCXLS-5469)
* Support I/O of Pivot Table Timeline slicer.(DOCXLS-6695)
* Add and manage Scenarios in What-If analysis.(DOCXLS-7905)
* Support for Pattern Fill when rendering to PDF/Image.(DOCXLS-8780)
* Support calculated page numbers in headers and footers.(DOCXLS-9370)
* Set color using various string formats: hex, RGB, and color names.(DOCXLS-9478)
* Support Image Sparkline formula.(DOCXLS-10009)
* Support automatic cell merging.(DOCXLS-10109)
* Support Cell Decoration styling.(DOCXLS-10566)
* Support option to include or exclude binding data.(DOCXLS-10645)
* Support Excel Table as data source for Pivot Tables.(DOCXLS-10763)
* Support for allow edit in cell, Chart Color Scheme, and Date-Time Picker options in Lossless I/O for SpreadJS..(DOCXLS-10997)
* Support new APIs to manage PivotTable.(DOCXLS-11101)
## Fixed
* The hidden data lables of chart are shown in the exported image file.(DOCXLS-6125,DOCXLS-11325)
* The Timeline Slicer is lost in the exported PDF file.(DOCXLS-10291)
* The rowcount is incorrect after calling IRange.Delete() method.(DOCXLS-10567)
* The data lables of chart are lost in the exported HTML file.(DOCXLS-10736)
* Some style is not expanded as expected after processing template.(DOCXLS-10897)
* Performance issue in template processing.(DOCXLS-10906)
* The fill color of data points in chart is incorrect in the exported Image file.(DOCXLS-11080)
* The size of the exported SJS file is much larger than the source file.(DOCXLS-11105)
* The layout of RadioButtonListCellType is different from SpreadJS.(DOCXLS-11191)
* Some cells are not expanded as expected after processing template.(DOCXLS-11282)
* The result of IWorksheet.RowCount is incorrect after setting it explicitly.(DOCXLS-11304).
* Upgraded the versions of certain dependencies to eliminate security vulnerabilities(DOCXLS-11562).
* Exception is thrown on processing particular template.(DOCXLS-11591)

## 7.2.5
## Fixed
* When copying array formulas and exporting them to Excel, the reference ranges change.(DOCXLS-11412)
* The form fields are incorrect in the exported PDF file after processing template.(DOCXLS-11416)
* Some conditional formats are lost in the exported Excel file.(DOCXLS-11419)
* Exception is thrown on opening a specific Excel file with a password.(DOCXLS-11445)
* The XMATCH function does not return correct value if lookup array contains empty cells.(DOCXLS-11448)
* Exception is thrown on loading SJS file contains null author in comment.(DOCXLS-11469)
* The autofit method does not work after setting wrap text to the cell.(DOCXLS-11478)
* The slicers are lost in the exported Excel file after processing template.(DOCXLS-11479)
* Spill error shows in array functions in the exported Excel file when calculation engine is off.(DOCXLS-11488)
* The conditional formatting is incorrect in the exported Excel file after processing template.(DOCXLS-11491)
* Exception is thrown on calling calculating method when custom formula returns CalcError.(DOCXLS-11498)
* Exception is thrown on processing template when filter condition contains negative number.(DOCXLS-11504)
* The hidden rows/columns are lost after calling IRange.copy() method.(DOCXLS-11506)
* Exception is thrown on processing template when template file contains specific chart.(DOCXLS-11531)
* Update the versions of some dependencies to eliminate security vulnerabilities.(DOCXLS-11562)
* Exception is thrown on processing template if the filtered field is empty.(DOCXLS-11567)
## 7.2.4
## Fixed
* OutOfMemoryException is thrown on exporting Excel file to SJS file.(DOCXLS-11029)
* The font name in specific cells has changed after loading the SSJSON file.(DOCXLS-11190)
* Some unused named styles are missing after loading the original SJS file.(DOCXLS-11218)
* Some data labels are missing in the exported image file.(DOCXLS-11220)
* Exception is thrown on opening an SJS file.(DOCXLS-11283)
* The chart is corrupted in the SJS file exported by DsExcel.(DOCXLS-11319)
* Emoji characters in sheet name gets corrupted in the exported SJS file.(DOCXLS-11336)
* The fill color is incorrect after adding a new row to the table in the exported Excel file.(DOCXLS-11344)
* The result of LET function is incorrect when parameter contains '#'.(DOCXLS-11364)
* The result of MAX function is incorrect in Template Language.(DOCXLS-11367)
* Exception is thrown on exporting PDF file when the legend position of chart is invalid.(DOCXLS-11373)
* The result of VLOOKUP function is incorrect when the parameter contains table formula.(DOCXLS-11379)
* The result of TAKE function is incorrect when the parameter contains negative value.(DOCXLS-11386)
* The result of SWITCH function is incorrect when it is used with other dynamic array formulas.(DOCXLS-11398)
* The result of MAXIFS function is incorrect when it is used in defined name.(DOCXLS-11401)
* The result of COUNTIFS function is incorrect.(DOCXLS-11402)
## 7.2.3
## Fixed
* The alignment is incorrect in the exported Excel file when overwrite fillmode is applied to cross table in template.(DOCXLS-10421)
* The invocation of the PrintManager consumes a lot of memory.(DOCXLS-10807)
* The plot area is not in the center in the exported image file.(DOCXLS-11084)
* Some resources are incorrect in the de-DE/de-CH/it-CH/fr-CH culture in the exported Excel file.(DOCXLS-11103)
* Exception is thrown on exporting SJS file after copying formula from another workbook.(DOCXLS-11111)
* Enhance the exception message when inserting column is invalid.(DOCXLS-11183)
* Exception is thrown on exporting to SJS file.(DOCXLS-11196)
* The conditional format is incorrect in the exported XLSX file after process template.(DOCXLS-11206)
* The reference of formula is changed in the exported XLSX file after process template.(DOCXLS-11207)
* The transparent of text in chart is lost in the exported PDF file.(DOCXLS-11215)
* The min/max date of chart axis is incorrect in the exported PDF file.(DOCXLS-11217)
* The result of CONCATENATE function is incorrect.(DOCXLS-11219)
* Enhance the error message when there are invalid content in template cells.(DOCXLS-11232)
* Some borders are missing in the exported PDF file.(DOCXLS-11262)
* The color of hyperlink is incorrect in the exported SJS file.(DOCXLS-11266)
## 7.2.2
## Fixed
* There are many unexpected diagonal lines in the exported PDF file.(DOCXLS-10988)
* Exception is thrown on exporting HTML file if some shape width is zero.(DOCXLS-11082)
* The exported Excel file is corrupted when the source data of the PivotChart contains CalcError.(DOCXLS-11102)
* The result of AVERAGEIF function is incorrect.(DOCXLS-11117)
* Exception is thrown on opening an SJS file that contains invalid quality factor.(DOCXLS-11118)
* The text that exceeds the boundaries of the shapes has not been cropped in the exported PDF file.(DOCXLS-11119)
* After setting the value to a cell range, the row count of the sheet becomes 1048576.(DOCXLS-11122)
* The sheet order is incorrect in the exported SJS file.(DOCXLS-11123)
* Performance issue on calling ProcessTemplate method.(DOCXLS-11124)
* Exception is thrown on opening an XLSX file that contains invalid DocPros data.(DOCXLS-11125)
* Exception is thrown on setting the orientation of PivotField as ColumnField.(DOCXLS-11129)
* Performance issue on deleting rows.(DOCXLS-11130)
* The result of EXACT function is different from Excel.(DOCXLS-11133)
* After copying the sheet, the cell displayed as #BUSY contains IMAGE function in the exported PDF file.(DOCXLS-11136)
* The result is 0 when the cell refers to the IMAGE function in the exported PDF file.(DOCXLS-11137)
* The text is overlapped in the exported PDF file when setting particular font.(DOCXLS-11155)
* The value of the SUMPRODUCT function is incorrect.(DOCXLS-11158)
## 7.2.1
## Fixed
* The macro button is missing after processing template file.(DOCXLS-10883)
* The XValues of chart are incorrect in the exported PDF file.(DOCXLS-10893)
* The static cell did not expand with the adjacent cell.(DOCXLS-10902)
* Performance issue on calling ToJson method.(DOCXLS-10903)
* The value and reference range of defined name become "#REF!" after processing template.(DOCXLS-10908)
* "_Report_0" is added to the end of the shape name after processing template.(DOCXLS-10931)
* There is a null fallback FontFamily in the exported HTML file.(DOCXLS-10969)
* Exception is thrown on setting formula of Calculated Fields in Switzerland culture locale.(DOCXLS-10976)
* There is a memory leak when using the ResultSet data source.(DOCXLS-10986)
* After setting TimeSpan to cell value, the number format changes.(DOCXLS-10990)
* The formula in table is lost after binding data.(DOCXLS-10994)
* Exception is thrown on copying sheet to another workbook when sheet contains column header.(DOCXLS-11005)
* The hidden rows are revealed when calling ITable.ConvertToRange method.(DOCXLS-11014)
* Exception is thrown on exporting to Excel file when workbook contains invalid threaded comment.(DOCXLS-11055)
* The Japanese currency symbol ¥ is not recognized correctly.(DOCXLS-11057)
* The result of getting the cell value is inconsistent.(DOCXLS-11059)
* The formula is lost after data binding.(DOCXLS-11067)
* The image is lost in the exported PDF file when image function contains special content.(DOCXLS-11068)
* The row height is lost after processing template.(DOCXLS-11073)
* Performance issue on adding lots of hyperlinks.(DOCXLS-11074)
* The cell value cannot be set correctly after setting default value.(DOCXLS-11078)
* Some resources are incorrect in the Deutsch(German) culture in the exported Excel file.(DOCXLS-11103)
* The warning message is incorrect when copying sheet that some references do not exist.(DOCXLS-11104)
* The column width is lost after processing template twice.(DOCXLS-11113)
## 7.2.0
## Added
* Support Asynchronous IMAGE function.(DOCXLS-6740)
* Support Label and Value Filter in PivotTable.(DOCXLS-7568)
* Support Distributed Alignment in PDF Export.(DOCXLS-7705) 
* Support Pixel-Based rendering in PDF and Image Export.(DOCXLS-8690)
* Support UID attribute when loading and saving Excel file.(DOCXLS-8756)
* Support filtering data from single/multiple data sources in Template Language.(DOCXLS-8835)
* Support display of filter button in column header for SJS and SSJSON export.(DOCXLS-9036)
* Support LINESPARKLINE/COLUMNSPARKLINE/WINLOSSSPARKLINE functions.(DOCXLS-9324)
* Support Goal Seek functionality.(DOCXLS-9389)
* Support customising Border Style in PDF Export.(DOCXLS-9624)
* Support for FromSJSJson API to load single SJS JSON file.(DOCXLS-9928)
## Fixed
* The image size is not same as the PDF file exported by SpreadJS.(DOCXLS-897)
* The exported XLSM file is corrupted when opening in MSExcel.(DOCXLS-5849)
* When a specific xlsx file is loaded and saved, the saved file is corrupted.(DOCXLS-6040)
* Wrapped text cannot be fully displayed in the exported PDF file.(DOCXLS-6762,DOCXLS-10267)
* The pagination results in the exported PDF file are inconsistent with SpreadJS.(DOCXLS-7128,DOCXLS-7897,DOCXLS-8743,DOCXLS-8827,DOCXLS-10278,DOCXLS-10334)
* The rotated circular shapes are distorted in the exported image files.(DOCXLS-8526)
* The line breaks in the exported PDF file are inconsistent with the PDF file exported by SpreadJS.(DOCXLS-8664,DOCXLS-8898,DOCXLS-9359,DOCXLS-10161,DOCXLS-10385)
* The size of the QR code in the exported PDF file is not same as SpreadJS.(DOCXLS-8801)
* The borders of some cells are lost in the exported Excel file after loading the SJS file.(DOCXLS-9244)
* The cell format is changed in the exported SJS file and Excel file after loading the original SJS file.(DOCXLS-9272,DOCXLS-9575,DOCXLS-9866,DOCXLS-10025,DOCXLS-10156)
* The cell format is changed in the exported SJS file after inserting specific rows.(DOCXLS-9341)
* The page margins in the exported PDF file are inconsistent with SpreadJS.(DOCXLS-9500)
* The seal is stretched in the exported PDF file.(DOCXLS-9588)
* The cell background image is inconsistent with SpreadJS in the exported PDF file.(DOCXLS-9611)
* The filtering status of PivotTable in the exported Excel file is changed.(DOCXLS-9613)
* The table border in the exported PDF file is thicker than the PDF exported by SpreadJS.(DOCXLS-9620)
* Some whitespace characters are missing in the exported PDF file.(DOCXLS-9722,DOCXLS-10365,DOCXLS-10888)
* The table style is changed in the exported SJS file.(DOCXLS-9876)
* The border style is changed in the exported SJS file.(DOCXLS-9993)
* Exception is thrown on refreshing the PivotTable.(DOCXLS-10239)
* The font size is changed in the exported SJS file after loading the original SJS file.(DOCXLS-10447)
* The conditional format is incorrect in the exported SJS file.(DOCXLS-10453)
* Incorrect value returned from array formula when calculation engine is off.(DOCXLS-10709)
* The table style in the exported PDF file is incorrect.(DOCXLS-10774)
* Some cell values are not exported in the Excel file when license is not set.(DOCXLS-10805)
* The result of the LOG function is incorrect.(DOCXLS-10822)
* Exception is thrown on setting the orientation of PivotField to column field and row field.(DOCXLS-10828)
* The result of the SEARCH function is incorrect.(DOCXLS-10859)

## 7.1.5
## Fixed
* The formula becomes dynamic array formula in the exported Excel file after processing the template.(DOCXLS-8816)
* The custom formatter is lost in the exported SSJSON file.(DOCXLS-10560)           
* The source data of the Pivot Table is not updated after data binding.(DOCXLS-10643)
* Some constant text is lost in the exported Excel file when setting FillMode as overwrite.(DOCXLS-10647)
* The applied range of conditional formatting is incorrect in the exported Excel file after processing the template.(DOCXLS-10648)
* The IndentLevel is reset to 0 after setting horizontal alignment as distributed.(DOCXLS-10657)
* Exception is thrown on exporting image file when pfb font exists.(DOCXLS-10714)
* The formula becomes dynamic array formula in the exported Excel file after processing the template.(DOCXLS-10715)
* The active sheet is changed in the exported Excel file after processing the template.(DOCXLS-10717)
* Exception is thrown on loading invalid Excel file generate by third-party component.(DOCXLS-10719)
* The print area is incorrect when it is relative reference.(DOCXLS-10720)
* Exception is thrown on exporing to SJS file after copying worksheet to another workbook.(DOCXLS-10724)
* The result of IRange.getFormula() changes from the origninal formula value when it contains absolute reference.(DOCXLS-10730)
* The custom data labels in chart are lost in the exported HTML file.(DOCXLS-10734)
* The row count and column count are incorrect in the exported Excel file after processing template when template contains comments.(DOCXLS-10738)
* The default value is lost in the exported SJS file.(DOCXLS-10739)
* Exception is thrown on saving particular Excel file to PDF file.(DOCXLS-10754)
* Exporting to SJS file twice, results in the size of the generated file not being the same.(DOCXLS-10758)
* Performance issue on exporting SJS and SSJSON files.(DOCXLS-10769)
* The row height is changed in the exported SSJSON file compared to the original one.(DOCXLS-10809)
## 7.1.4
## Fixed
* Merged cells are unmerged in the exported Excel file when template cell have G=R property.(DOCXLS-10484)
* The serise line of chart changed after loading the SSJSON file then exporting to SJS file.(DOCXLS-10488)
* The formula is lost after loading the particular SJS file.(DOCXLS-10495)
* The font weight is lost after loading the particular SSJSON file.(DOCXLS-10498)
* Performance issues on opening a particular Excel file which contains lots of external workbooks.(DOCXLS-10506)
* Exception is thrown on expoting Excel file after deleting particular sheet.(DOCXLS-10510)
* The customized properties of celltype are lost in the exported SSJSON file.(DOCXLS-10553)
* Exception is thrown on setting dynamic array formula to data validation.(DOCXLS-10571)
* Exception is thrown on calling Workbook.toJson() method when the original SSJSON file contains invalid comment data.(DOCXLS-10589)
* The result of TEXTJOIN function is incorrect if the parameter contains CalcError.(DOCXLS-10617)
* The text is incorrect in the exported PDF file when the font is "symbol".(DOCXLS-10639)
## 7.1.3
## Fixed
* The column style is lost after calling processTemplate method in the exported XLSX file.(DOCXLS-10338)
* There is a performance issue when saving specific Excel files to an HTML file.(DOCXLS-10342)
* Enhance the exception message when deleting range is not allowed.(DOCXLS-10364)
* Exception is thrown on filling the specific template file.(DOCXLS-10379)
* Row heights are changed after opening an XLSX file then export to SSJSON file.(DOCXLS-10386)
* Exception is thrown on deleting row after loading some particular SJS file.(DOCXLS-10389)
* Constant cells are expand along with the template cell.(DOCXLS-10417)
* Exception is thrown on calling processTemplate method after loading some particular file.(DOCXLS-10418)
* Some zero value appera after process template in the export XLSX file.(DOCXLS-10422)
* Exception is thrown on exporting to SJS file.(DOCXLS-10428)
* Exception is thrown on calling IWorksheets.CopyAfter method.(DOCXLS-10448)
* Performance issue on loading particular XLSX file.(DOCXLS-10457)
* The value of allowMove property in chart changed in the exported SJS file.(DOCXLS-10461)
* The text changes to vertical text in the exported PDF file.(DOCXLS-10482)
## 7.1.2
## Fixed
* The formula result is incorrect using API to get the cell value.(DOCXLS-10106)
* Performance issue on calculating workbook containing dynamic array formulas.(DOCXLS-10174)
* The number format is incorrect when using Chinese version of Excel to open the Excel file exported in US culture.(DOCXLS-10234)
* Rich text is missing in the exported PDF file.(DOCXLS-10247)
* Copying the sheet takes a long time when sheet contains lots of pictures.(DOCXLS-10279)
* The print info is changed in the exported SSJSON file.(DOCXLS-10280)
* An exception is thrown on saving Excel file.(DOCXLS-10284)
* Chart data lost after copying worksheet if EnableCalculation is false.(DOCXLS-10289)
* Performance issue on exporting specific SSJSON file.(DOCXLS-10292)
* German text changes to English language when refreshing the Pivot Table.(DOCXLS-10293)
* Exception is thrown on applying sorting via FontColorSortField.(DOCXLS-10314)
* Exception is thrown on loading particular Excel file.(DOCXLS-10320)
* Exception is thrown on processing the template when it contains cross worksheet formula.(DOCXLS-10326)
* Exception is thrown on processing the template when it contains hidden sheets.(DOCXLS-10329)
* All rows are hidden in the exported SJS file.(DOCXLS-10335)
* The comments are lost in the exported Excel file after processing template.(DOCXLS-10336)
## 7.1.1
## Fixed
* Sheet scrolling buttons are not working when hiding sheet and exporting to Excel file.(DOCXLS-8480)
* The text direction is different from Excel in the exported PDF file.(DOCXLS-10091)
* The formula result is incorrect when reading the formula cell value and saving back to Excel.(DOCXLS-10106)
* The formula result is incorrect with Japanese full-width parameters.(DOCXLS-10115)
* Exception is thrown on saving an Excel after copying range with formula.(DOCXLS-10116)
* Exception is thrown on saving Excel to PDF file with Espanol culture.(DOCXLS-10117)
* Some cell values have an additional decimal place compared to Excel in the exported PDF file.(DOCXLS-10157)
* The axis lable doesn't renders correctly in the PDF file.(DOCXLS-10165)
* Exception is thrown on opening an Excel file.(DOCXLS-10173)
* The exported Excel file size is very large and time-consuming.(DOCXLS-10182)
* Exception is thrown on saving to the Excel file.(DOCXLS-10187)
* The row count is changed in the exported SSJSON file.(DOCXLS-10211)
* Exception is thrown on loading an SJS file.(DOCXLS-10236)
* Long cell value is lost in the exported PDF file when using PrintManager.(DOCXLS-10249)
## 7.1.0

## Added
* Enhance Template Language with better performance.(DOCXLS-5722)
* Ignore errors in range.(DOCXLS-8310)
* Support for interrupting execution of ProcessTemplate method.(DOCXLS-8396)
* Use custom fonts via font streams on PDF Export.(DOCXLS-8432)
* Template Language supports OverwriteWithFormat under Classic Mode.(DOCXLS-8515)
* Support lossless .sjs/ssjson I/O of GanttSheet.(DOCXLS-8629)
* Support CalculationMode options.(DOCXLS-8750)
* Search cells with tags.(DOCXLS-8759)
* Support multi column sorting in Template Language.(DOCXLS-8818)
* Support custom sort order in Template Language.(DOCXLS-8819)
* Get and set cell background images.(DOCXLS-8829)
* Support table reference in cross workbook formulas.(DOCXLS-8859)
* Export Barcodes as Pictures in Excel files.(DOCXLS-9278)
* Support lossless .sjs/ssjson I/O of ReportSheet.(DOCXLS-9625)
## Fixed
* Exception is thrown on calling ProcessTemplate method when data source is not found.(DOCXLS-5052)
* The grouped result is incorrect after calling ProcessTemplate method when using JSON data source.(DOCXLS-6476)
* The result of SUMIFS formula is incorrect after processing template file.(DOCXLS-6596)
* Exception is thrown on calling ProcessTemplate method when data field is numeric.(DOCXLS-8098)
* The result in the report workbook is incorrect after processing the template file.(DOCXLS-8489)
* The result of SUM formula is incorrect after processing the template file.(DOCXLS-8514)
* The merge area is incorrect after processing the template file.(DOCXLS-8656)
* The result of DIV formula is incorrect after processing the template file.(DOCXLS-8669)
* The cell background color is not expanded as expected in the exported report workbook.(DOCXLS-8709)
* The result of MIN and MAX formula is incorrect after processing template file.(DOCXLS-8786)
* The calculated result is incorrect in the exported report workbook after processing template file.(DOCXLS-8834)
* The formula is changed in the exported report workbook after processing template file.(DOCXLS-9002)
* Exception is thrown on calling ProcessTemplate method.(DOCXLS-9023)
* Performance issue on processing template when setting G=R.(DOCXLS-9041)
* The apply range of conditional format is changed in the exported workbook.(DOCXLS-9046)
* The chart in the exported PDF file is incorrect.(DOCXLS-9579)
* In the exported PDF file, there are some Chinese characters that appear as garbled or scrambled.(DOCXLS-9778)
* The ISort.apply() method didn't work as intended.(DOCXLS-9853)
* Performance issue on processing template containing merged cell.(DOCXLS-9961)
* The border is lost in the exported XLSX file.(DOCXLS-10015)
* Image is lost in the exported XLSX file when PaginationMode is true.(DOCXLS-10020)
* The row/column header caption in Pivot Table is lost in the exported XLSX file.(DOCXLS-10044)
* Exception is thrown on loading particular SSJSON file.(DOCXLS-10047)
* Exception is thrown on loading particular SJS file.(DOCXLS-10084)
* The result of the DATEVALUE formula is incorrect.(DOCXLS-10085)
* Performance issue on processing template when there're too many sheets.(DOCXLS-10088)
* The table formula is not expanded correctly when calling ProcessTemplate method.(DOCXLS-10090)
* Exception is thrown on loading worksheet from JSON containing non-existent external references.(DOCXLS-10092)

## 7.0.5
## Fixed
* Some content in the hyperlink is escaped after loading the SJS file.(DOCXLS-9926)
* The table range is not updated after data binding.(DOCXLS-9946)
* Exception is thrown on expoting PDF file when workbook contains specific cell button.(DOCXLS-9947)
* Exception is thrown on processing template.(DOCXLS-9949)
* Exception is thrown on loading Excel file.(DOCXLS-9951)
* Exception is thrown on loading SJS file when some table formula is empty.(DOCXLS-9955)
* Performance issue on accessing or setting named style.(DOCXLS-9957)
* Performance issue on getting or setting value when workbook contains dynamic array formulas.(DOCXLS-9994)
* The KeepTogether result is in correct in the generated Excel file.(DOCXLS-9997)
* Specific svg image is lost in the exported PNG file.(DOCXLS-9999)
* Exception is thrown on processing template.(DOCXLS-10002)
* Grouped shapes are clipped off in the exported image file.(DOCXLS-10003)
* The font size of axis label in chart is incorrect in the exported image file.(DOCXLS-10004)
* Exception is thrown on loading particular SSJSON file.(DOCXLS-10010)
* Exception is thrown on loading then exporting SJS file when it contains particular tags.(DOCXLS-10013)
## 7.0.4
## Fixed
* The checkstatus of the checkbox is incorrect in the exported PDF file.(DOCXLS-4050)
* The cell value is incorect in the exported PDF file when value is an array.(DOCXLS-9790)
* Exception is thrown on exporting SJS file when data validation contains invalid formulas.(DOCXLS-9791)
* Exception is thrown on getting cell value when the cell refers to non-existing worksheet.(DOCXLS-9794)
* The formula changed after inserting a column when formula contains non-existing worksheet.(DOCXLS-9796)
* The image in page setup is lost in the exported SSJSON file.(DOCXLS-9806)
* Exception is thrown on opening a particular Excel file.(DOCXLS-9823)
* The shape is not copied using IRange.Copy() method.(DOCXLS-9854)
* Merge area changed in the exported Excel file.(DOCXLS-9862)
* The copying option in Excel is failed in the exported Excel file.(DOCXLS-9914)
## 7.0.3
## Fixed
* The captions of Pivot Table item are reset after refreshing it.(DOCXLS-9160)
* Exception is thrown on exporting workbook to PDF file.(DOCXLS-9586)
* Performance issue when copying cells to another workbook.(DOCXLS-9623)
* The border in the exported HTML file is inconsistent with the display in SpreadJS.(DOCXLS-9627)
* The processing time of the template has significantly increased after the upgrade.(DOCXLS-9697)
* Exception is thrown on trying to get the DataValidation Formula.(DOCXLS-9708)
* Integers have an extra decimal place in the exported file after template processing.(DOCXLS-9715)
* Exception is thrown on saving to SJS file after loading the original one.(DOCXLS-9723)
* The border is lost in the exported SJS file after calling IRange.Copy() method.(DOCXLS-9725)
* Exception is thrown on loading specific Excel file.(DOCXLS-9731)
* DefaultValue property doesn't work correctly in a certain case.(DOCXLS-9741)
## 7.0.2
## Fixed
* The formula changes after inserting columns if the referenced worksheet does not exist.(DOCXLS-7267)
* After importing SJS and SSJSON files exported by DsExcel into the SpreadJS designer, the charts appear differently.(DOCXLS-9362)
* The cell value is incorrect after loading the SJS file exported by SpreadJS.(DOCXLS-9365)
* The exported Excel file is corrupted after loading the SSJSON file then exporting.(DOCXLS-9453)
* Exception is thrown on calling GetAccurateRangeBoundary method for particular Excel if the calc engine is turned off.(DOCXLS-9502)
* Exception is thrown on exporting PDF file when workbook contains specific chart.(DOCXLS-9532)
* Exception is thrown on loading the specific SSJSON file exported by SpreadJS.(DOCXLS-9572)
* The layout of cell background image is incorrect in the exported PDF file.(DOCXLS-9585)
* Exception is thrown on opening SSJSON file when calc engine is turned off.(DOCXLS-9593)
* The number format is changed in the exported SJS file.(DOCXLS-9605)
* The font size is incorrect in the exported PDF file.(DOCXLS-9606)
* Exception is thrown on refreshing Pivot Table.(DOCXLS-9607)
* The column width is not adjusted after calling IRange.Autofit() method.(DOCXLS-9681)
* The result of COUNTA(UNIQUE()) function is incorrect.(DOCXLS-9683)
* Exception is thrown on adding sheet after setting formula.(DOCXLS-9686)
## 7.0.1
## Fixed
* It takes more than 5 minutes to export specific SSJSON file to PDF file.(DOCXLS-8989)
* Unexpected diagonal border is exported in the XLSX file.(DOCXLS-9173)
* The result of AGGREGATE formula is incorrect.(DOCXLS-9287)
* The font style of shape is incorrect in the exported PDF file.(DOCXLS-9300)
* The icon of conditional format in the exported HTML file is different from Excel and SpreadJS.(DOCXLS-9302)
* Some back color of cells change to black in the exported PDF file.(DOCXLS-9313)
* Exception is thrown on loading a specific XLSX file.(DOCXLS-9335)
* Exception is thrown on loading a specific XLSX file.(DOCXLS-9339)
* Exception is thrown on loading a specific SSJSON file if the file contains invalid chart data.(DOCXLS-9340)
* Performance issue while calculating workbook.(DOCXLS-9360)
* The conditional formatting styles are lost after loading an SSJSON file then exporting to SJS file.(DOCXLS-9364)
* The call of method PrintManager.paginate() takes too much time.(DOCXLS-9387)
* When the special xlsx with hyperlinks is loaded and saved, the saved xlsx is corrupted.(DOCXLS-9395)
* Some cells that were not locked became locked after loading an SJS file then exporting to another SJS file.(DOCXLS-9401)
* The returned value of fromJson does not contain formula error.(DOCXLS-9402)
* Exception is thrown on calling Workbook.ToJson() method.(DOCXLS-9406)
* Exception is thrown on opening an SSJSON file contains invalid formulas.(DOCXLS-9416)
* The result of VDB function is different from Excel.(DOCXLS-9427)
* The formula is not updated after calling IWorksheet.FromJson().(DOCXLS-9433)
* The text is incorrect in the exported PDF file.(DOCXLS-9448)
* The center header of page setup can not be set using setText method.(DOCXLS-9479)
* Exception is thrown on copying worksheet to another workbook.(DOCXLS-9510)
## 7.0.0

## Important note for users of the GrapeCity Documents for Excel, Java Edition
* com.grapecity.documents:gcexcel is being renamed and will be continued to be maintained under the new names com.mescius.documents:dsexcel. These new packages provide the same functionality, ensure future enhancements, and are backwards compatible with GrapeCity Documents for Excel. Please update your references to avoid any possible future interruptions. Your existing subscriptions will continue to work with the new packages.

* com.grapecity.documents:gcexcel 已经更名为 com.grapecitysoft.documents:gcexcel。
* com.grapecitysoft.documents:gcexcel 与 com.grapecity.documents:gcexcel 的功能保持一致，并且保持兼容，未来我们将持续对 com.grapecitysoft.documents:gcexcel 添加新功能并进行维护。请将您的引用从 com.grapecity.documents:gcexcel 改为 com.grapecitysoft.documents:gcexcel 以避免任何可能的问题。
* com.grapecitysoft.documents:gcexcel 与 com.grapecity.documents:gcexcel 使用相同的授权策略，如果在您的升级中，遇到任何授权问题，可以发送邮件至 info.xa@grapecity.com 联系我们获取帮助。

## Breaking Changes from the Previous Release
* The return value of Parent property of IDataLabel interface has been changed from IPoint to Object.

## Added
* Support Formatting for Trendline Equations in Charts (and export).(DOCXLS-5769)
* Async User-Defined Function Support.(DOCXLS-6971)
* Export Excel to HTML with Inline CSS option.(DOCXLS-7433)
* Improvements in Grouping for OptionButton Controls.(DOCXLS-7784)
* [Template Language]Maintain Image aspect ratio.(DOCXLS-8216)
* Support Acroform Creation with GcExcel API.(DOCXLS-8341)
* [SpreadJS Integration]Support for password in the protected sheet.(DOCXLS-8421)
* Support exporting of Funnel Charts to PDF.(DOCXLS-8570)
* [SpreadJS integration]Support for Mask style.(DOCXLS-8618)
* Support for IRange.DefaultValue.(DOCXLS-8651)
* [SpreadJS integration]Support for cell.altText property.(DOCXLS-8652)
* [Template Language]Support shapes and images in pagination mode.(DOCXLS-8694)
* Support exporting of smooth lines in chart to PDF.(DOCXLS-8710)
* Set first page number to 'Auto' in Page Setup.(DOCXLS-8734)
* Specify columns to quote on exporting to CSV.(DOCXLS-8795)

## Fixed
* Formula linked with radio button disappeared in the exported Excel file.(DOCXLS-7688)
* The celltype is changed in the exported SSJSON file.(DOCXLS-8348)

## 6.2.5
## Fixed
* The size of the QR code is not as expected.(DOCXLS-8801)
* After calling IWorksheet.copyAfter() method, an error occurs when opening the exported Excel file.(DOCXLS-8901)
* Performance issue when calling IRange.Copy() method many times.(DOCXLS-9157)
* The icon of conditional format in the exported PDF file is different from Excel and SpreadJS.(DOCXLS-9177)
* Exception is thrown on copying sheet across Excel files when sheet contains custom sheet views.(DOCXLS-9179)
* The selected area is a whole row, the used range would be changed in the exported SJS file.(DOCXLS-9182)
* The picture cannot be fully displayed in the exported PDF file.(DOCXLS-9194)
* When a worksheet is copied twice to another workbook, the defined names are not copied correctly.(DOCXLS-9217)
* The number format of NOW function is incorrect under Korean culture.(DOCXLS-9218)
* The formula has been changed in the exported Excel file after loading the original SSJSON file.(DOCXLS-9240)
* Cell headers style is incorrect in the exported SSJSON file after loading the original SSJSON file.(DOCXLS-9243)
* The exported Excel file gets corrupted with particular sheet name and Print Area detail.(DOCXLS-9246)
* Exception occurred during data binding.(DOCXLS-9254)
## 6.2.4
## Fixed
* Chart is not output correctly in PDF file when combo chart of scatter charts have no data.(DOCXLS-9027)
* Exception is thrown on Workbook.processTemplate() method.(DOCXLS-9053)
* The cell formatting is incorret in the copied worksheet.(DOCXLS-9058)
* Exception is thrown on loading the particular SSJSON file contains ComboBox celltype.(DOCXLS-9063)
* Exception is thrown on calculating the function which refers to a non-exist worksheet.(DOCXLS-9106)
* Exception is thrown on saving workbook to the Excel file.(DOCXLS-9153)
* The exported XLSX file is corrupted when the hyperlink contains quote.(DOCXLS-9174)
* Inconsistent results when processing template with multiple threads.(DOCXLS-9180)
## 6.2.3
## Fixed
* There are unexpected lines in the exported Excel file after processing Template file.(DOCXLS-8921)
* DataValidation Error Alert doesn't preserves after saving the Workbook.(DOCXLS-8923)
* InvalidOperationException is thrown on specifying multiple ranges for a chart data source.(DOCXLS-8962)
* The calculated result is inconsistent between the exported PDF and Excel files.(DOCXLS-8987)
* Threaded comments are not visible in Finnish locale.(DOCXLS-8995)
* Size of Chart Y Axis is incorrect in the exported Excel file.(DOCXLS-8996)
* There are unexpected characters in the exported PDF file.(DOCXLS-9020)
## 6.2.2
## Fixed
* The result of XLOOKUP function is incorrect.(DOCXLS-8691)
* The text wrapping position is inconsistent with SpreadJS.(DOCXLS-8784)
* The picture is incorrect in the exported PDF file when picture has crop and rotation setting.(DOCXLS-8791)
* The center header of printinfo is lost in the exported SSJSON file.(DOCXLS-8800)
* Exception is thrown on loading the SSJSON file that contains invalid chart data.(DOCXLS-8806)
* Exception is thrown on rendering PDF file when workbook contains particular chart.(DOCXLS-8857)
* Performance issue on setting cell value.(DOCXLS-8858)
* The interior of cell is incorrect in the exported PDF file.(DOCXLS-8891)
## 6.2.1
## Fixed
* The legend of deleted series in chart are displayed in the exported Excel file.(DOCXLS-8452)
* Doughnut charts is lost in the exported PDF file.(DOCXLS-8530)
* The wrap position is incorrect when text contains "-" charactor in the exported PDF file.(DOCXLS-8535)
* The result of the formula in the defined name is not updated correctly.(DOCXLS-8563)
* The font is incorrect in the exported PDF file.(DOCXLS-8575)
* The program gets stuck when processing template.(DOCXLS-8578)
* The sparklines is incorrect in the exported PDF file.(DOCXLS-8586)
* The program gets stuck when processing template.(DOCXLS-8598)
* Exception is thrown on exporting PDF file when workbook contains sparklines.(DOCXLS-8603)
* Exception is thrown on evaluating LAMBDA function.(DOCXLS-8620)
* Exception is thrown on loading particular SSJSON file.(DOCXLS-8633)
* Exception is thrown on loading the SSJSON made by customer.(DOCXLS-8654)
* Some shape, picture and lines are lost in the exported PDF file.(DOCXLS-8685)
* The cell text is not displayed completely in the exported PDF file.(DOCXLS-8686)
* Exception is thrown on exporting PDF file when page footer contains special font.(DOCXLS-8692)
* Exception is thrown on opening particular Excel file contains unsupported vml drawings.(DOCXLS-8711)
* Diagonal lines disappeared in the exported Excel file.(DOCXLS-8713)
* Incorrect cell value is returned when EnableCalculation = False.(DOCXLS-8725)
* Performance downgradation on calling calculate method.(DOCXLS-8728)
* Sheet name is lost after loading the SSJSON file.(DOCXLS-8735)
* The pagination result is incorrect in the exported PDF file.(DOCXLS-8744)
* Exception is thrown on loading SSJSON file after disable the calculate engine.(DOCXLS-8754)
## 6.2.0
## Added
* Set Vertical text direction in a Shape and Chart.(DOCXLS-6334)
* Alignment options for Shape Text.(DOCXLS-6793)
* Support header reference.(DOCXLS-8055)
* Support for SpreadJS .sjs file format.(DOCXLS-8004)
## Fixed
* The result of XLOOKUP formula is incorrect.(DOCXLS-8542)

## 6.1.4
## Fixed
* The result of XIRR func is incorrect.(DOCXLS-8450)
* InvalidFormulaException is thrown on opening an Excel file with XLOOKUP formula which refers to external Table formula.(DOCXLS-8478)
* The result of reappling filter is incorrect after setting data source.(DOCXLS-8481)
* The result of COUNTIFS func is incorrect.(DOCXLS-8483)
* The result of dynamic array formula is incorrect when disabling calculation engine the enabling.(DOCXLS-8503)
* StackOverFlowException is thrown on processing template when there are circular reference.(DOCXLS-8511)
* The size of the QR code is incorrect in the exported PDF file.(DOCXLS-8513)
* The multiple fields in the same template cell do not take effect in the exported report workbook.(DOCXLS-8514)
* Performance issue on calculating SUMPRODUCT function.(DOCXLS-8516)
* The data bar is lost in the exported PDF file.(DOCXLS-8524)
* The result of WEEKNUM formula is incorrect.(DOCXLS-8531)
* The shape disappears in the exported image after setting rotation.(DOCXLS-8540)
* InvalidCastException is thrown on exporting PDF file.(DOCXLS-8542)
## 6.1.3
## Fixed
* Performance issue on processing Template file.(DOCXLS-7932)
* Exception is thrown on calling IRange.ungroup() method.(DOCXLS-8355)
* The result of DATEDIF formula is incorrect in JP culture.(DOCXLS-8378)
* Some cell content is lost after processing Template file.(DOCXLS-8385)
* Exception is thrown on copying worksheet contains form controls.(DOCXLS-8412)
* After refreshing the Pivot Table, the exported SSJSON file size becomes very large.(DOCXLS-8443)
* Performance issue on formula calculation.(DOCXLS-8455)
* The row height is incorrect when cell has center across selection in the exported PDF file.(DOCXLS-8465)
* The row height is incorrect when meger cell has not set wrap text.(DOCXLS-8473)
* The transparency setting of shape is lost in the exported Excel file.(DOCXLS-8476)
## 6.1.2
## Fixed
* Performance downgradation on using PrintManager to exporting PDF file comparing to old version.(DOCXLS-7599)
* The column count is incorrect in the exported PDF file when adding a shape with rotation.(DOCXLS-7724)
* The result of cross workbook formula is incorrect in the export Excel file.(DOCXLS-8244)
* The calculated item is not savd correctly when some fields are not visible.(DOCXLS-8249)
* Performance issue on calling Workbook.calculate() method.(DOCXLS-8260)
* The indent of the checkbox list is missing in the exported PDF file.(DOCXLS-8264)
* KeyNotFoundException is thrown on exporting PDF file when workbook contains specific chart.(DOCXLS-8268)
* The sheet content is incorrect in the exported PDF file.(DOCXLS-8269)
* The texts containing line breaks are not output correctly in the PDF file.(DOCXLS-8273)
* The custom number format is not applied to cell which has rich text in the exported PDF file.(DOCXLS-8278)
* The length of certain shape is incorrect in the exported PDF file.(DOCXLS-8281)
* Chart data is lost after merging workbooks containing chart and sheet reference formula.(DOCXLS-8282)
* The conditional formattings of cells with fills are not output to PDF file correctly.(DOCXLS-8290)
* The result of SUMIF function is incorrect.(DOCXLS-8306)
* The stacked series is lost when category axis is date type in the exported PDF file.(DOCXLS-8308)
* Exception is thrown on loading some irregular Excel file.(DOCXLS-8336)
* Exception is thrown on calling IRange.clear() method.(DOCXLS-8339)
* The exported Excel report file is incorrect after processing template.(DOCXLS-8346)
* OutOfMemoryError is thrown on refreshing pivot table.(DOCXLS-8349)
* Exception is thrown on saving PDF file when the original workbook contains abnormal data.(DOCXLS-8356)
* The result of IRange.getUsedRange() method is incorrect.(DOCXLS-8358)
## 6.1.1
## Fixed
* Performance issue on calling IRange.merge() method and exporting to Excel file.(DOCXLS-6001)
* Exception is thrown on loading particular SSJSON file.(DOCXLS-8071)
* The result is incorrect of data binding when the binding path contains the index of array.(DOCXLS-8074)
* The tick marks units of charts in the exported PDF file are smaller than ones in Excel file.(DOCXLS-8089)
* The margin setting is incorrect in the exported Excel file after loading the SSJSON file.(DOCXLS-8092)
* Performance issue on calling calculate method when workbook contains LOOKUP functions.(DOCXLS-8097)
* Filter is lost in the exported SSJSON file.(DOCXLS-8115)
* The column width of row header is lost in the exported SSJSON file.(DOCXLS-8121)
* The exported Excel file is corrupted when the original file contains group shape.(DOCXLS-8125)
* The series is incorrect in the exported Excel file when the chart contains references to defined names of other worksheet.(DOCXLS-8139)
* The result of the formula is incorrect when formula contains defined name which has volatile function.(DOCXLS-8142)
* The result of OR function is incorrect.(DOCXLS-8144)
* Unexpected fore color style is in the exported SSJSON file.(DOCXLS-8149)
* Performance issue on loading sjs file.(DOCXLS-8152)
* The number format is changed in the exported SSJSON file.(DOCXLS-8156)
* Some content is lost on loading the particular SSJSON file.(DOCXLS-8190)
* The result of VLOOKUP function is incorrect.(DOCXLS-8193)
* The cellpadding setting takes no effect on the exported PDF file when cell contains celltype.(DOCXLS-8207)
* The fore color is incorrect when cell has number format with colors in the exported SSJSON file.(DOCXLS-8218)
* Exception is thrown on loading the particular Excel file then calling ToJson method.(DOCXLS-8219)
* Performance issue on calling ToJson method.(DOCXLS-8221)
* The drop-down box is lost in the exported SSJSON file.(DOCXLS-8224)
* Exception is thrown on loading particular Excel file which contains invalid content.(DOCXLS-8234)
* Exception is thrown on exporting Excel file after loading the SSJSON file contains abnormal chart title.(DOCXLS-8237)
## 6.1.0
## Added
* Export options in ToImage() method.(DOCXLS-5481)
* Copy/move multiple sheets at once.(DOCXLS-6332)
* Support for XLTX File Format.(DOCXLS-6343)
* Support Form Controls on SpreadJS JSON I/O.(DOCXLS-6872)
* Support for allowResize property on SpreadJS JSON I/O.(DOCXLS-6875)
## Fixed
* Exception is thrown on adding calculated items on Pivot Field.(DOCXLS-7800)                 
* The performance of Workbook.fromJson() method is bad when JSON contains stripped style in big tables.(DOCXLS-7892)
* The calculating performance is bad when sheet contains SUMPRODUCT formula.(DOCXLS-7902)
* The axis position of chart is incorrect in the exported PDF file.(DOCXLS-7914)
* Exception is thrown on loading particular Excel file with chart.(DOCXLS-7922)
* The calculated formula of table column is lost in the exported SSJSON file.(DOCXLS-7930)
* 3D chart is missing in the exported PDF file.(DOCXLS-7984)
* The result of MATCH function is incorrect.(DOCXLS-8013)
* The result of COUNT function is incorrect.(DOCXLS-8023)
* The cell value is not overflowed in the exported PDF file when cell value type is text and has customer number format.(DOCXLS-8025)
* The result of some functions is incorrect when they refer to 3D reference.(DOCXLS-8048)
* Exception is thrown on exporting PDF file when combo chart's category axis has null value in the workbook.(DOCXLS-8063)

## 6.0.6
## Fixed
* Image size is not adjusted with the column width when setting BestFitColumns as true when exporting PDF file.(DOCXLS-7782)                 
* NullReferenceException is thrown on calling the PrintManager.Paginate() method.(DOCXLS-7808)
* The text is not wrapped in the exported PDF file.(DOCXLS-7828)
* Cell notes lost after saving workbook to a new Excel file.(DOCXLS-7831)
* Data binding does not support special symbols.(DOCXLS-7837)
* Table filters are not working properly after fetching Table JSON.(DOCXLS-7843)
* Table and Slicers are cleared after fetching the table JSON.(DOCXLS-7844)
* Exception is thrown on setting particular calculated field formula for Pivot Table.(DOCXLS-7854)
* The application gets hanged on loading the XLSX file.(DOCXLS-7856)
* The time validation info is wrong in the SSJSON file exported by DsExcel.(DOCXLS-7861)
* Bad performance of IRange.RowHeight setter.(DOCXLS-7866)
* When the cell with incorrect formula is copied to another workbook, the workbook is corrupted.(DOCXLS-7871)
* Defined names are lost in the exported SSJSON file.(DOCXLS-7873)
* AutoParse does not take affect when using data binding.(DOCXLS-7874)
* Connector shapes are lost in the exported PDF file.(DOCXLS-7884)
* The text style is lost in the copied shape when shape has linked text.(DOCXLS-7885)
* The chart plot area becomes smaller in the exported PDF file when chart is a grouped item.(DOCXLS-7889)
## 6.0.5
## Fixed
* Exception is thrown on calling calculated or save method when workbook contains dynamic array formulas.(DOCXLS-7404)
* The error type of formulas are different from Excel.(DOCXLS-7593)
* The condition format result in the exported PDF file is different Excel.(DOCXLS-7631)
* The table style is incorrect in the exported Excel file after processing template.(DOCXLS-7632)
* Exception is thrown on calling fromJson method.(DOCXLS-7641)
* Lambda function does not return correct result when getting values.(DOCXLS-7649)
* Exception is thrown on saving workbook when it contains corrupted ActiveX controls.(DOCXLS-7677)
* Sorting with blank rows is not worked correctly.(DOCXLS-7681)
* Exception is thrown on exporting PDF file through PrintManager when workbook contains hidden column.(DOCXLS-7716)
* Exception is thrown on calling IPivotTable.refresh() method.(DOCXLS-7723)
* The text layout is different from the Excel file and the exported PDF file.(DOCXLS-7738)
* The character is not escaped in the exported JSON file.(DOCXLS-7746)
* The link color of hyperlink is lost in the exported JSON file.(DOCXLS-7755)
* Exception is thrown on loading particular JSON file exported by DsExcel.(DOCXLS-7756)
* The column width is inconsistent in different pages in the exported PDF file.(DOCXLS-7757)
## 6.0.4
## Added
* Option to control if export shared formula in SSJSON file.(DOCXLS-7477)
## Fixed
* Exception is thrown on loading particular SSJSON file contains circular reference.(DOCXLS-7377)
* The calculated result is incorrect from Excel.(DOCXLS-7472)
* The result of IFNA function is incorrect from Excel.(DOCXLS-7537)
* The celltype in copied range is not cleared when copying range from other worksheet.(DOCXLS-7563)
* The column which width is 0 is invisible in the SpreadJS designer.(DOCXLS-7565)
* The shape is lost in the exported Excel and PDF file after setting the shape position.(DOCXLS-7598)
* Performance downgradation on using PrintManager to exporting PDF file comparing to old version.(DOCXLS-7599)
* The transparency effect in the shape text is lost in the exported PDF file.(DOCXLS-7604)
* The rotation of the shape is counterclockwise in the exported PDF file.(DOCXLS-7606)
* The protection setting is lost in the exported SSJSON file.(DOCXLS-7614)
* Stack overflow exception is thrown on calling calculate() method.(DOCXLS-7617)
* The exported Excel file is corrupted when original file contains invalid Pivot Table data.(DOCXLS-7633)
* IllegalArguments exception is thrown on loading particular Excel file.(DOCXLS-7634)
## 6.0.3
## Fixed
* The calculation result is incorrect when the formula is array formula.(DOCXLS-4598)             
* The number format changes from "General" to "Custom" after loading and exporting SSJSON file.(DOCXLS-6854)
* The height of plot area and the tick lables are incorrect in the exported PDF file.(DOCXLS-7076)
* Exception is thrown on inserting rows when worksheet contains data validation.(DOCXLS-7364)
* Exception is thrown on loading particular SSJSON file contains unexpected binding path content.(DOCXLS-7365)
* Exception is thrown on loading then saving to another Excel file.(DOCXLS-7376)
* The style of chart's trendline is incorrect in the exported PDF file.(DOCXLS-7381)
* The arrow mark is missed in the exported PDF file.(DOCXLS-7383)
* Exception is thrown on loading particular SSJSON file.(DOCXLS-7386)
* The table formula is lost when loading particular SSJSON file.(DOCXLS-7401)
* The checkbox format is different from the SpreadJS designer and exported PDF file.(DOCXLS-7412)
* Exception is thrown on loading Excel file contains external reference.(DOCXLS-7429)
* Unexpected cell content is shown in the exported PDF file.(DOCXLS-7430)
* Exception is thrown on acessing group shape's children using iterator.(DOCXLS-7434)
* The "colHeaderVisible" setting is lost after loading the original SSJSON file then exporting to another.(DOCXLS-7435)
* Exception is thrown on adding validation which contains INDIRECT function.(DOCXLS-7442)
* Exception is thrown on loading irregular Excel file then exporting to SSJSON file.(DOCXLS-7444)
* The transparency setting of the text in the shape is lost in the exported PDF file.(DOCXLS-7471)
* The underline is lost in the exported PDF file.(DOCXLS-7476)
* Exception is thrown on setting Hyperlink funtion in a copied worksheet.(DOCXLS-7510)
## 6.0.2
## Fixed
* The marker size, data label and series line color chart in incorrect in the exported PDF file.(DOCXLS-7229)
* Some row heights are incorrect after processing template.(DOCXLS-7268)
* The linked cell text color becomes white after loading the SSJSON file then exporting to SSJSON file.(DOCXLS-7293)
* Exception is thrown on saving Excel file when the file name has only one character.(DOCXLS-7314)
* The interface Workbook.SetLicenseFile is lost.(DOCXLS-7315)
## 6.0.1
## Fixed
* Performance issue when using table bindings with 500,000 rows data.(DOCXLS-6781)
* The exported Excel file is corrupted after processing template.(DOCXLS-6826)
* The data validation lost after loading the SSJSON file when the applied range is whole column.(DOCXLS-7082)
* Exception is thrown on exporting PDF file when using data binding.(DOCXLS-7094)
* PageSetup.Zoom does not take effect on barcode in the exported PDF file.(DOCXLS-7095)
* "Swiss 721 BT" font is not shown correctly in the exported PDF file.(DOCXLS-7102)
* Print titles settings are lost in the generated report after processing template.(DOCXLS-7126)
* The formula result is different from Excel.(DOCXLS-7168)
* Exception is thrown on loading particular XLSX file contains invalid data.(DOCXLS-7171)
* Exception is thrown on loading SSJSON file then calling Workbook.ToJson() method when JSON contains dupicated named styles.(DOCXLS-7174)
* Exception is thrown on calling processing template when data field name contains number.(DOCXLS-7194)
* PageSetup.CustomPaperSize method did not working.(DOCXLS-7195)
* The DateTimePicker cell button cannot be deleted when deleting row.(DOCXLS-7199)
* The cell text is incorrect after loading the Excel file contains rich text.(DOCXLS-7200)
* The applied range of conditional formatting is incorrect after coping cell.(DOCXLS-7214)
* Exception is thrown on rendering HTML when workbook contains Pivot Table.(DOCXLS-7219)
* Exception is thrown on processing template when sheet contains defined names.(DOCXLS-7220)
* The calculation result of VARP function differs from Excel.(DOCXLS-7221)
* Exception is thrown on loading particular XLSX file contains invalid data.(DOCXLS-7223)
## 6.0.0
## Breaking changes
* Changed XlsxOpenOptions.DoNotAutofitAfterOpened to XlsxOpenOptions.DoNotAutoFitAfterOpened(F in caps).(DOCXLS-6428)
## Added
* Support for new Lambda function including help functions.(DOCXLS-3667)
* Paginated Spreadsheet reports new enhancements.(DOCXLS-3829)
* New text and array manipulation Excel functions.(DOCXLS-5711)
* Support 'RowColumnStates' in JSON I/O.(DOCXLS-5828)
* Add shape text with range reference or defined name.(DOCXLS-5829)
* Add shape/picture to cell/cell range using direct method.(DOCXLS-6182)
* Add option to control Auto Fit.(DOCXLS-6228)
* Support cross workbook formula - 'externalReference' in JSON I/O.(DOCXLS-6243)
* Get used range in selected area.(DOCXLS-6330)
* Support page number and page count for groups in template pagination.(DOCXLS-6374)
* Keep original template and return the report workbook instance.(DOCXLS-6417)
* Process specific template worksheets in template language.(DOCXLS-6418)
* Excel workbook size optimization through save options.(DOCXLS-6441)
* Range intersection, union and offset.(DOCXLS-6487)
* DsExcel Java now targets JDK 8.(DOCXLS-6650)
* Support for 'allSheetsListVisible' field in JSON I/O.(DOCXLS-6696)

## 5.2.5
## Fixed
* The column width of exported PDF is inconsistent with SpreadJS.(DOCXLS-6857)
* IndexOutOfRangeException is thrown on calling Calculate method twice for a workbook.(DOCXLS-6873)
* Performance issue in opening Excel file and saving to JSON when workbook contains lots of comments.(DOCXLS-6886)
* The exported report is incorrect when sheet name is bind to the datafield with DataSet as datasource.(DOCXLS-6889)
* The layout is incorrect on call the IPivotTable.refresh() method.(DOCXLS-6890)
* Field name is missed in the exported report when using the Range and Group feature of template at the same time.(DOCXLS-6892)
* The border is lost in the exported SSJSON file.(DOCXLS-6896)	
* Loading the specific SSJSON file, some node would be discarded.(DOCXLS-6902)
* Template is not processed correctly.(DOCXLS-6909)
* Table formula is not updated when using template.(DOCXLS-6916)
* The exported SSJSON file is incorrect comparing with the original one.(DOCXLS-6970)
* Performance issue with Add function in dynamic array.(DOCXLS-6979)
* Exception is thrown on loading Excel file which contains irregualr content.(DOCXLS-6981)
* Exception is thrown on loading particular SSJSON file with Pivot Table.(DOCXLS-6982)
* InvalidOperationException is thrown on getting cell value twice from a workbook.(DOCXLS-6988)
* The number format is missed when data source contains date value.(DOCXLS-6996)
* The row height is incorrect after calling autofit	method.(DOCXLS-7005)
* The cell value is incorrect when workbook contains cross workbook reference.(DOCXLS-7008)
* Exception is thrown on loading particular SSJSON file contains conditional format with entire column reference.(DOCXLS-7024)
* The SUMIF function result is incorrect.(DOCXLS-7032)
## 5.2.4
## Fixed
* Incorrect results with iterative calculations after loading the particular Excel file.(DOCXLS-6444)              
* When the url wrapped in a cell is exported to pdf, it breaks in the wrong position.(DOCXLS-6547)
* Performance issue when setting the hidden property of the row.(DOCXLS-6588)
* CharSet of the QR code is missed in the exported pdf file.(DOCXLS-6765)
* NullPointerException is thrown on setting cell values in worksheet which contains iterative calculation.(DOCXLS-6779)
* Exception is thrown on calculating MIDB function in CJK culture.(DOCXLS-6791)
* The hidden rows are unhidden after loading the SSJSON file.(DOCXLS-6806)
* Bad performance when loading SSJSON file contains lots of named styles.(DOCXLS-6821)
* Print area is lost after loading the ssjson file.(DOCXLS-6824)
* VBA project digital signature is lost in the exported Excel file.(DOCXLS-6831)
* The style is incorrect after calling the copy() method.(DOCXLS-6853)
* The column width of exported PDF is inconsistent with SpreadJS.(DOCXLS-6857)
* The path in external link is not encoded correctly.(DOCXLS-6862)
* The program would not exit if the repeat area is large when exporting PDF file.(DOCXLS-6866)
* The count of Pivot items is incorrect.(DOCXLS-6867)
* Exception is thrown on loading particular SSJSON file contains #SPILL!.(DOCXLS-6874)
* IndexOutOfRangeException is thrown on copying a sheet from one workbook to another.(DOCXLS-6878)
* Merged cells unmerged after processing the template.(DOCXLS-6882)
* When cells with long wrapped texts are autofitted and exported to PDF file, the row heights are not adjusted.(DOCXLS-6883)
* The position of barcode is incorrect in the exported PDF file.(DOCXLS-6884)
* Print range and margin settings are lost in the exported report when pagination mode is on.(DOCXLS-6888)
* Exception is thrown on exporting SSJSON file when workbook contains particular defined names.(DOCXLS-6901)
## 5.2.3
## Fixed
* Performance issue when exporting the Excel file to JSON when workbook contains lots of table with striped styles.(DOCXLS-5242)
* The result of IPivotItem.Visible is incorrect when the option "select multiple items" is checked.(DOCXLS-6385)
* The result of INDEX function is not spilled correctly.(DOCXLS-6483)
* When Excel file including "Add-ins" is loaded and saved, "Add-ins" is lost.(DOCXLS-6535)
* The cell value is not formatted correctly in the exported HTML file.(DOCXLS-6555)
* An OutOfMemoryError occurs when exporting specific Excel file to image.(DOCXLS-6577)
* There would be #Ref error when the sheet name starts with number.(DOCXLS-6580)
* The file is bigger after exported.(DOCXLS-6595)
* Performance issue when calculating lots of formulas.(DOCXLS-6597)
* The result of get hidden columns is incorrect when the Excel is exported by other product.(DOCXLS-6639)
* ArrayIndexOutOfBoundsException is thrown on exporting Excel file.(DOCXLS-6640)
* The EUDC font could not be displayed correctly in the exported PDF file.(DOCXLS-6670)
* The picture is lost in the exported PDF file.(DOCXLS-6671)
* Json parse error is thrown when the row count or column count is string in conditional format.(DOCXLS-6679)
* InvalidFormulaException is thrown on refreshing data in the PivotTable.(DOCXLS-6697)
* The position of the bent connector's arrow is incorrect in the exported image file.(DOCXLS-6699)
* "This Year" setting in conditional format is not supported after loading JSON file.(DOCXLS-6712)
* Floating object is lost in the exported JSON file.(DOCXLS-6721)
* The layout of databar is incorrect in the exported HTML file.(DOCXLS-6725)
* Performance issue when exporting to PDF file when condition format contains whole column reference.(DOCXLS-6730)
* The cell type on table column is lost after loading the JSON file.(DOCXLS-6741)
* Exception is thrown on loading particular Excel file.(DOCXLS-6748)
## 5.2.2
## Fixed
* NumberFormatException is thrown on loading the particular excel file.(DOCXLS-6134)           
* Bad Performance when exporting PDF.(DOCXLS-6428)
* Exception is thrown on deleting the series collection of chart.(DOCXLS-6442)
* Font color changed when importing SSJSON file to workbook and saving back to Excel file.(DOCXLS-6455)
* The result of INDEX function is not spilled correctly.(DOCXLS-6483)
* ArgumentOutOfRange exception is thrown on copying a sheet from one workbook to another.(DOCXLS-6485)
* NumberFormatException is thrown on loading the particular excel file.(DOCXLS-6489)
* Z-order of pictures is changed when loading SSJSON file and exporting to Excel file.(DOCXLS-6496)
* The result of IRange.Text is incorrect when showzero is false.(DOCXLS-6525)
* Invalid key data exception is thrown when adding jakarta.json dependency.(DOCXLS-6545)
* Exception is thrown on loading SSJSON file when conditional format contains gradient color.(DOCXLS-6556)
* The exported Excel file is corrupted when calling IRange.sort() method.(DOCXLS-6565)
* Twill pattern of conditional format would be lost after loading SSJSON file.(DOCXLS-6578)
* Grouped shapes do not move with cells.(DOCXLS-6584)
* Exception is thrown on loading particular Excel file.(DOCXLS-6636)
* Exception is thrown on exporting to PDF file when workbook contains charts.(DOCXLS-6638)
## 5.2.1
## Fixed
* The results of IWorksheet.FreezeColumn/FreezeRow are incorrect.(DOCXLS-6299)
* Exception is thrown on processing template if the data source name is number and Japanese.(DOCXLS-6300)
* The result of Search function is incorrect when it contains the wild char "\*".(DOCXLS-6361)
* Poor performance and high memory footprint when loading particular Excel file.(DOCXLS-6362)
* The exported Excel file is corrupted when workbook contains bubble chart.(DOCXLS-6372)
* The slicer in the exported ssjson file is different from the original JSON file.(DOCXLS-6384)
* The layout and the data field value of pivot table in exported XLSX file is incorrect.(DOCXLS-6406)
* The data field result of pivot table is not filterd in the exported XLSX file.(DOCXLS-6412)
* The result of INames.Contains() is incorrect.(DOCXLS-6429)
* Exception is thrown on loading SSJSON file when font name contains unexpected characters.(DOCXLS-6436)
* Repeated loading and saving of an existing file corrupts the saved Excel file.(DOCXLS-6470)
## 5.2.0
## Added
* New API to add Excel Form Controls.(DOCXLS-2532)
* Support for CASCADESPARKLINE formula.(DOCXLS-3562)
* Define Paginated Templates in spreadsheets (with DsExcel Templates).(DOCXLS-3829)
* Support for Chart Data Table.(DOCXLS-4438)
* Support for LET function.(DOCXLS-5185)
* GetPivotData Formula supports spilled data.(DOCXLS-5390)
* Add Calculated Item in Pivot Table.(DOCXLS-5437)
* Get accurate range boundary.(DOCXLS-5473)
* Support JSON as DataSource in data binding.(DOCXLS-5499)
* IsVolatile property support in Custom function.(DOCXLS-5714)
* Debug mode in DsExcel Templates.(DOCXLS-5715)
* Debug better with additional details in InvalidFormulaException.(DOCXLS-5744)
* Support for adding and rendering SVG image.(DOCXLS-5817)

## 5.1.5
## Fixed
* The performance of opening particular Excel file need to be optimized.(DOCXLS-6038) 
* It takes a long time to execute calculating when workbook contains AverageIf formula.(DOCXLS-6055)
* The calculation result of MIDB formula is incorrect in CJK culture.(DOCXLS-6112)
* The IMPOWER formula result is incorrect.(DOCXLS-6173)
* Exception is thrown on calling Workbook.toJson() method when workbook contains external reference.(DOCXLS-6177)
* Exception is thrown on loading particular SSJSON.(DOCXLS-6189)
* The calculation result of SUMIFS formula is inconsistent with Excel.(DOCXLS-6197)
* Exception is thrown on calling IRange.fromJson() method.(DOCXLS-6199)
* The calculation result of AGGREGATE formula is incorrect.(DOCXLS-6209)
* The allowDynamicArray setting is lost in the exported SSJSON file.(DOCXLS-6210)
* One static row is missing after processing the template using DsExcel.(DOCXLS-6212)
* The ignore blank setting of DataValidation is changed after loading the SSJSON file.(DOCXLS-6217)
## 5.1.4
## Fixed
* The rowheight is incorrect after copying worksheet.(DOCXLS-5845)
* The result of SEARCHB func is different from Excel in Chinese culture.(DOCXLS-5994)
* The result of autoFit is bad in the exported PDF file.(DOCXLS-6028)
* Some borders are missing in the exported image after loading particular JSON file.(DOCXLS-6097)
* Exception is thrown on drawing Pivot Table to PDF file.(DOCXLS-6110)
* IllegalArgumentException is thrown on exporting image when some fonts are missed in system.(DOCXLS-6123)
* The texts in the merged cell near the page boundary are not output to PDF file.(DOCXLS-6137)
* ArgumentException is thrown on saving Excel file when workbook contains some special characters.(DOCXLS-6138)
* Invalid Argument exception is thrown on loading csv from URL stream.(DOCXLS-6142)
* The last cell of column is incorrect when a column is deleted.(DOCXLS-6144)
* The rows are hidden in the exported JSON file.(DOCXLS-6153)
* NullReferenceException is thrown on saving to Excecl file.(DOCXLS-6155)
* IllegalArgumentException is thrown on loading particular JSON file.(DOCXLS-6158)
* The texts in merged cell are not exported in the PDF file.(DOCXLS-6159)
* Exception is thrown on saving to PDF file if workbook contains corrupted image.(DOCXLS-6164)
* ArgumentException is thrown on loading a particular JSON file contains different slicer with same name.(DOCXLS-6166)
* IndexOutOfRangeException is thrown on getting some cell's value.(DOCXLS-6168)
* The exported Excel file is corrupted after refreshing the Pivot Table.(DOCXLS-6174)
## 5.1.3
## Fixed
* The color of cell font is incorrect in exported JSON file.(DOCXLS-5945）
* Platform dependent characters are incorrect in the exported PDF file.(DOCXLS-5971）
* NullReferenceException is thrown on exporting PDF file.(DOCXLS-5996）
* The performance of inserting lots of pictures is bad.(DOCXLS-6009)
* The value of formula N is incorrect.(DOCXLS-6010)
* The exported xlsm file is corrupted after deleting some sheet which contains external reference.(DOCXLS-6012)
* Setting IWorksheetView.ScrollColumn/ScrollRow properties on a worksheet which has freeze pane corrupts the saved Exel file.(DOCXLS-6026）
* Series lines in a chart are not exported to PDF file.(DOCXLS-6030)
* The zoom scale of the worksheet changes from 90% to 100%.(DOCXLS-6042)
* The exported Excel file is corrupted when the rows are deleted in the worksheet which has freeze pane.(DOCXLS-6054)
* Using copy method to copy range will throw System.NullReferenceException if pasteType value is NumberFormats.(DOCXLS-6056)
* Exception is thrown on opening the Excel file exported by DsExcel.(DOCXLS-6066)
* The cell value is disordered after calling IRange.Clear() method when workbook contains phonetic text.(DOCXLS-6083)
## 5.1.2
## Fixed
* The chart in the exported ssjson file is different from the ssjson exported by SpreadJS.(DOCXLS-5753）
* Data binding does not support multi-layer level.(DOCXLS-5875）
* The result of IRange.toImage() is incorrect.(DOCXLS-5884）
* The exported xlsx file is corrupted when copying range from one sheet to another.(DOCXLS-5888）
* [Template Language]The generated xlsx file is corrupted when template file has multiple sheets.(DOCXLS-5889)
* Some cell values are #REF! after using data binding.(DOCXLS-5913）
* The size of exported chart image is smaller than the original chart in MSExcel.(DOCXLS-5938)
* The exported xlsx file is corrupted when comment contains special characters.(DOCXLS-5942）
* The cell value is #REF! when calculating SUMIF function.(DOCXLS-5944)
* The position of chart's tick labels are incorrect in the exported xlsx file.(DOCXLS-5946)
* The rowcount is incorrect after loading CSV file from stream.(DOCXLS-5951)
* Performance degradation comparing with old versions.(DOCXLS-5952)
* The cell content is incorrect when the alignment is distributed and indent in exported PDF file.(DOCXLS-5953)
* The exported image is corrupted when the range is entire column/row.(DOCXLS-5990)
## 5.1.1
## Fixed
* [Template language]The layout is incorrect when rendering list of data with dynamic object in both right and down direction.(DOCXLS-4727)
* When shapes with text are exported to the PDF file, shapes are not the same as Excel file.(DOCXLS-5723)
* Illegal access warning occurs when using JDK9 or above.(DOCXLS-5752)
* ArgumentOutOfRangeException or NullReferenceException is thrown on using DsExcel API in multi-threaded environment when there's TEXT function in workbook.(DOCXLS-5784)
* The picture's position is incorrect in the exporting PDF file when some column is hidden.(DOCXLS-5798)
* Some borders are lost in the exported Excel file after loading SSJSON file.(DOCXLS-5813)
* The fontsize of cell is incorrect after loading the JSON file.(DOCXLS-5818)
* NullPointerException is thrown on loading JSON file.(DOCXLS-5821)
* The bottom border is lost in the exported PDF file.(DOCXLS-5838)
* The tick labels' orientation of chart is changed in the exported Excel file.(DOCXLS-5840)
* [Template language]Chart is incorrect after processing template in the exported Excel file.(DOCXLS-5844)
* #REF! error occurs when calculating OFFSET function.(DOCXLS-5847)
* NullPointerException is thrown on calling IPivotTable.refresh() method.(DOCXLS-5850)
* The first space of cell text is not preserved in the exported Excel file if the cell value is calculated by formula.(DOCXLS-5851)
* The result of COUNT function is incorrect.(DOCXLS-5859)
* The exported Excel file is corrupted after processing template.(DOCXLS-5876)
* The program would be stuck when exporting to JSON when there's filter in workbook.(DOCXLS-5879)
* ArrayIndexOutOfBoundsException is thrown on calling Workbook.toJson() method.(DOCXLS-5883)
* The exported Excel file is corrupted after loading specific xlsx file containing chart.(DOCXLS-5887)
## 5.1.0
## Added
* [Pivot Table] Specify 'ShowValuesAs' Option for 'Values' Field.(DOCXLS-738)
* Add Numbers Fit Mode of overflow/mask.(DOCXLS-3696)
* Modify Password of Excel Documents.(DOCXLS-4076)
* [Pivot Table] Support for Calculated Field.(DOCXLS-4211)
* Support JSON data source for template language.(DOCXLS-4359)
* Support for "Show #N/A as an empty cell" in Chart.(DOCXLS-4442)
* Convert Table to Range.(DOCXLS-4601)
* Support CELL function.(DOCXLS-4645)
* CSV Custom Parser.(DOCXLS-5162)
* Support for Pivot table views (JSON I/O).(DOCXLS-5221)
* Import Data Function to Import Table, Range, or Worksheet from Excel Files.(DOCXLS-5276)
* Support for TableSheet (JSON I/O).(DOCXLS-5487)
## Fixed
* Converting the imported Excel to Json breaks all formulas in it.(DOCXLS-3840)
* [Template language]Out of memory exception is thrown on using template in multi-thread.(DOCXLS-4449)
* [Template language]The layout is incorrect in the exported xlsx file when using template with cross table.(DOCXLS-4756)
* [Template language]The result of sum function is incorrect in the exported xlsx file.(DOCXLS-4827)
* [Template language]The cell values are not expected in the exported Excel file.(DOCXLS-4888)
* [Template language]The IF formula is lost in the exported xlsx file.(DOCXLS-5224)
* [Template language]The formula is not expanded correctly in exported xlsx file.(DOCXLS-5239)
* IndexOutOfRangeException is thrown on exporting radar chart to PDF file.(DOCXLS-5262)
* [Template language]The result of sum function is incorrect in the exported xlsx file.(DOCXLS-5485)
* The legend of series is not shown in the exported PDF file if the datapoints of the series are empty.(DOCXLS-5569)
* InvalidFormulaException is thrown on creating table when there is "'" in the table column's name.(DOCXLS-5690)
* Extra characters are exported in data validation in the exported JSON file.(DOCXLS-5734)
* NullReferenceException is thrown on setting series formula which contains bubble size info.(DOCXLS-5751)
* The layout of checkbox list is incorrect in the exported PDF file.(DOCXLS-5754)
## 5.0.5
## Fixed
* Bad performance when calculating formulas in a specific Excel file.(DOCXLS-5501)
* The position of the frozen pane is changed after JSON I/O.(DOCXLS-5525)
* The layout of checkbox list in the exported pdf file is different from SpreadJS designer.(DOCXLS-5544)
* Performance is bad when calling method toJSON with particular Excel file.(DOCXLS-5564)
* When the chart is exported to PDF, the pattern fill of the legend symbol is not correct.(DOCXLS-5567)
* The layout of the list object in the exporting pdf is not correct.(DOCXLS-5590)
* The checkbox is unchecked when the cell value is set.(DOCXLS-5624)
* Open csv file with double quote, the records would be lost.(DOCXLS-5639)
* ParseException is thrown on opening particular excel file with threaed comment.(DOCXLS-5643)
* Exception is thrown on setting Workbook.Culture property when workbook contains Hyperlink named style.(DOCXLS-5649)
* Exception is throws when opening loading specific JSON file.(DOCXLS-5656) 
* Exception is thrown on calling Workbook.SetLicenseKey method in multi-thread.(DOCXLS-5660)
* The result of IRange.LastColumn is not correct in some Excel file.(DOCXLS-5668)
* The plot area of chart is not updated in the exported pdf file when enabling/disabling calculation.(DOCXLS-5682)
* Custom named style are not exported to JSON or Excel file if it is not used in any range.(DOCXLS-5683)
* Font "Noto Sans JP" is not shown correctly in the exported pdf.(DOCXLS-5701)
* Exception is thrown on opening an Excel file.(DOCXLS-5704)
* Exception is thrown on opening an Excel file with ActiveX control.(DOCXLS-5706)
* NullPointerException is thrown on calling Calculate method after loading JSON file.(DOCXLS-5709)
## 5.0.4
## Fixed
* [Template Language] Cell style is incorrect after calling IWorkbook.ProcessTemplate().(DOCXLS-5206)
* IndexOutOfRangeException is thrown on exporting pdf file when sheet contains hidden columns.(DOCXLS-5263)
* Column width is changed when importing JSON file exported by SpreadJS.(DOCXLS-5314)
* IndexOutOfBoundsException is thrown on exporting JSON file for specific Excel file.(DOCXLS-5377)
* Exception is thrown on saving an Excel file exported by DsExcel.(DOCXLS-5393)
* The exported Excel file is corrupted when it contains specific linked objects.(DOCXLS-5402)
* The exported Excel file is corrupted when it contains dynamic array formulas.(DOCXLS-5409)
* NullReferenceException is thrown on opening some excel files with waterfall chart.(DOCXLS-5418)
* Font in shape is changed in exported Excel file.(DOCXLS-5422)
* The linked picture does not change along with the refered range.(DOCXLS-5439)
* The tab of active sheet is not displayed in the tab strip.(DOCXLS-5441)
* The formula value is incorrect in the exported Excel file.(DOCXLS-5444)
* The exported Excel file is corrupted when copying range with dynamic array formula.(DOCXLS-5459)
* The row height changes after inserting rows.(DOCXLS-5465)
* Format settings are not copied when copying entire rows contains merge area.(DOCXLS-5472)
* Formula result is different from Excel in Chinese culture.(DOCXLS-5488)
* NullPointerException is thrown on calling toJson method for user specific excel.(DOCXLS-5523)
* The result of function "=NOW()" is not precise when time goes by.(DOCXLS-5524)
* When the BarChart is exported to PDF, the order of the legends is reversed.(DOCXLS-5536)
* XMLStreamException is thrown on saving Excel to JSON file.(DOCXLS-5540)
* An error occurred while copying sheet with threaded comment.(DOCXLS-5546)
* Binding result is incorrect when using bigdecimal as data source.(DOCXLS-5547)
* Particular headers and footers are not output to PDF correctly.(DOCXLS-5554)
* The style is incorrect in the exported PDF file when copying sheet with conditional format.(DOCXLS-5557)
* Exception of index out of range is thrown on exporting PDF file if the workbook contains specific conditional format.(DOCXLS-5588)
* LET function changed in the exported Excel file.(DOCXLS-5593)
* InvalidFormulaException is thrown on opening a workbook with specific charts.(DOCXLS-5596)
* The result of formula calculation is different with Excel.(DOCXLS-5598)
* Exception is thrown on exporting chart to image.(DOCXLS-5626)
## 5.0.3
## Breaking changes
* Support option to control whether treat BigDecimal as double for cell value.(DOCXLS-5364)
## Added
* Add ability to add filter without condition using IRange.AutoFilter().(DOCXLS-2435)
* Add options to control if serialize workbook to JSON without worksheets.(DOCXLS-5357)
## Fixed
* Formulas are lost in exported Excel file.(DOCXLS-5371)
* ToJson method generates invalid json file when Excel file having multi-line comments.(DOCXLS-5384)
* Cell value is incorrect when evaluating formula using DsExcel.(DOCXLS-5385)
* The cell style is incorrect in exported Excel file after loading the JSON exported by SpreadJS.(DOCXLS-5389)
* Exception is thrown on calling workbook.toJson() method.(DOCXLS-5398)
* Performance degradation when evaluating values from the particular Excel file.(DOCXLS-5400)
* The formula result of SUBSTITUTE is incorrect.(DOCXLS-5408) 
## 5.0.2
## Fixed
* Default column width changes when exporting to Excel file.(DOCXLS-5161)
* The method IRange.Autofit() would not take effect when adjacent cell is set wrapping.(DOCXLS-5264)
* Exception is thrown on saving Excel file when the formula2 in data validation is empty.(DOCXLS-5266)
* Exception is thrown on saving PDF file when workbook contains specific chart.(DOCXLS-5274)
* InvalidFormulaException is thrown on opening particular Excel file contains special characters in table column.(DOCXLS-5277)
* OutOfMemory exception is thrown on exporting the Excel file with shape object to HTML.(DOCXLS-5278)
* It takes a long time on calling Workbook.ToJson() method.(DOCXLS-5302)
* Some KR demo resources are missing.(DOCXLS-5317)
* Exception is thrown on loading the JSON file exported by DsExcel.(DOCXLS-5321)
* Exception is thrown on exporting PDF file when workbook contains specific Pivit Table.(DOCXLS-5325)
* StackOverFlow exception is thrown on opening specific Excel file.(DOCXLS-5328)
* Shape border is changed in the exported HTML file when resaving it.(DOCXLS-5335)
* Temporary file is created after opening the Excel file using DsExcel API.(DOCXLS-5396)
## 5.0.1
## Fixed
* The alignment style is lost in the exported JSON file.(DOCXLS-5198)
* The formula is incorrect in the exported Excel file when using template language to generate report.(DOCXLS-5205)
* Unexpected borders are shown in exported image.(DOCXLS-5209)
* FormatRows setting of worksheet protection is lost in the exported Excel file.(DOCXLS-5213)
* System.AggregateException is thrown on saving Excel file to multiple PDF files in parallel.(DOCXLS-5222)
* Hidden rows are shown in the exported JSON file.(DOCXLS-5225)
* IndexOutOfRangeException is thrown on calling 'Workbook.Calculate()' method.(DOCXLS-5227)
* The cell color is incorrect in the exported Excel file.(DOCXLS-5229)
* The QRCode which has connection symbol is not output to the PDF file correctly.(DOCXLS-5230)
* The radar chart is incorrect in the exported Excel file.(DOCXLS-5238)
* Unexpected borders are output in the exported PDF file if the cell is merged.(DOCXLS-5248)
* NullPointerException is thrown on calling 'Workbook.ToJson()' method.(DOCXLS-5253)
* The exported Excel file is corrupted when copying a sheet with name length is 31.(DOCXLS-5256)
## 5.0.0
## Added
* Support for Linked Picture.(DOCXLS-625)
* Support for Table expandBoundRows API.(DOCXLS-3447)
* Support for Excel threaded comments.(DOCXLS-3989)
* Enhance license system.(DOCXLS-4181)
* Import data function.(DOCXLS-4758)
* Support for workbook views.(DOCXLS-4760)
* New Formula2 property to set Dynamic Array Formula.(DOCXLS-4949)
* Support for GETPIVOTDATA Function.(DOCXLS-5200)
## Fixed
* The result of NOW() function is incorrect.(DOCXLS-3721)
* There is a large gap on the top of the cell in exported PDF file.(DOCXLS-4218)
* NullPointerException is thrown on calling 'Workbook.Open()' method.(DOCXLS-4886)
* The cell style is lost in the exported JSON file when compared with the original JSON file.(DOCXLS-4999)
* The pattern fill settings of charts are not rendered in exported PDF file.(DOCXLS-5071)
* The text with "\n" in charts is shifted in exported PDF file.(DOCXLS-5072)
* The cell values are incorrect in exported CSV file.(DOCXLS-5153)
* The cell styles are incorrect in exported Excel file after loading the JSON file.(DOCXLS-5154)
* The exported Excel file is corrupted when copying a worksheet from another worksheet.(DOCXLS-5155)
* The result of 'TEXT' function is incorrect.(DOCXLS-5173)
* ArrayIndexOutOfBoundsException is thrown on calling 'Workbook.ToJson()' method.(DOCXLS-5196)
## 4.2.6
## Fixed
* The cell formula and value are lost in exported JSON file.(DOCXLS-5089)
* It takes too much time on calling method 'Workbook.ProcessTemplate()'.(DOCXLS-5094)
* Exception is thrown on calling 'Workbook.ToJson()' method.(DOCXLS-5098)
* The cell style is incorrect in the exported Excel file.(DOCXLS-5099)
* Exception is thrown on calling 'Workbook.FromJson()' method.(DOCXLS-5112)
* The result of 'Round' formula is incorrect in German culture.(DOCXLS-5117)
* IlleagalArguement and NumberFormat exception would be thrown on opening particular excel file using DsExcel.(DOCXLS-5118)
* The formula in exported JSON by DsExcel is not the same as original JSON.(DOCXLS-5122)
* The floating object is lost in the exported JSON file.(DOCXLS-5125)
* Exception is thrown on opening particular excel file.(DOCXLS-5129)
* Exception is thrown on calling method 'Workbook.ToJson()' when workbook contains PivotTable.(DOCXLS-5140)
* Theme color is incorrect in the exported Excel file.(DOCXLS-5144)
* The version is incorrect in the exported JSON file.(DOCXLS-5145)
* Exception is thrown on calling method 'Workbook.FromJson()' when JSON contains conditional format.(DOCXLS-5149)
## 4.2.5
## Fixed
* It takes too much time on calling method 'GetUsedRange'.(DOCXLS-4552)
* It takes too much time on copying range with formulas.(DOCXLS-4884)
* Exception is thrown on calling IRange.FromJson.(DOCXLS-4922)
* The formula is not correct after cutting range to another position.(DOCXLS-4979)
* The exported Excel file is corrupted after copying a range from one workbook to another.(DOCXLS-4985)
* The number format is changed after setting cell's value.(DOCXLS-4987)
* Exception is thrown on calling 'ToJson' method in Mac environement.(DOCXLS-4997)
* Macros are not preserved when an xlsm file with OLE objects is loaded and saved.(DOCXLS-5027)
* The exported Excel file is corrupted after copying a worksheet from one workbook to another.(DOCXLS-5043)
* OLE Objects are broken in exported Excel file.(DOCXLS-5047)
* NumberFormatException is thrown on opening an Excel file which contains Pivot Table.(DOCXLS-5051)
* The filter result is incorrect in exported JSON file.(DOCXLS-5059)
* Extra characters are added to the comments in exported Excel file.(DOCXLS-5060)
* The rows that should be hidden are shown in exported Excel file.(DOCXLS-5061)
* Exception is thrown on calling 'workbook.calculate()' method.(DOCXLS-5079)
* It takes too much time on calling method 'IWorksheet.FromJson()'.(DOCXLS-5084)
## 4.2.4
## Fixed                              
* IGraphic.LockAspectRatio didn't take effect when adding picture or changing picture size.(DOCXLS-4175)
* An error would be thrown when loading the JSON file exported by DsExcel in SpreadJS designer.(DOCXLS-4778)
* Exception is thrown on opening excel file exported by YARG using DsExcel.(DOCXLS-4874)
* The exported JSON could not be validated.(DOCXLS-4900)
* InvalidFormulaException is thrown on opening an Excel file.(DOCXLS-4901)
* Exception is thrown on copying range.(DOCXLS-4903)
* Exception is thrown on loading JSON file exported by DsExcel.(DOCXLS-4908)
* Formula is lost after loading JSON and exporting to Excel file.(DOCXLS-4924)
* (Duplicate)InvalidFormulaException is thrown on opening an Excel file.(DOCXLS-4947)
* The comments are lost in exported JSON file.(DOCXLS-4953)
* ConcurrentModificationException is thrown on copying workbook range.(DOCXLS-4965)
* Exception is thrown on saving Excel file using DsExcel.(DOCXLS-4988)
## 4.2.3
## Enhanced
* Internal compatibility additions with DsExcel .NET
## 4.2.2
## Enhanced
* Support XlsmOpenOptions and XlsmSaveOptions.(DOCXLS-4747)
* Upgrade the pdfbox version to 2.0.24.(DOCXLS-4787)
## Fixed                              
* Exception is thrown on copying a worksheet to another workbook.(DOCXLS-4737)
* The checkbox is lost in the exported PDF file.(DOCXLS-4744)
* The cell value is incorrect after using setValue method to set the cell value.(DOCXLS-4781)
* Print Area is different from the original JSON file in the exported JSON.(DOCXLS-4789)
* Exception is thrown on reading the JSON file.(DOCXLS-4798)
* The exported Excel file is corrupted after deleting rows using DsExcel API.(DOCXLS-4799)
* It takes too much time on calling the ToJson method.(DOCXLS-4803)
* The cell value is incorrect on using setValue method when the value is a BigDecimal.(DOCXLS-4806)
* Exception is thrown on calling the toJson method when the margin of the printinfo is null.(DOCXLS-4808)
* The formula result "=EDATE(NOW(),1)" calculated by DsExcel is different from Excel.(DOCXLS-4816)
* NumberFormatException is thrown on loading a particular excel file with user.language=de.(DOCXLS-4817)
* It takes too much time on opening a workbook.(DOCXLS-4818)
* The grouped result using DsExcel is different from Excel.(DOCXLS-4823)
* Exception is thrown on calling the IRange.Group()and workbook.Save() methods.(DOCXLS-4824)
* NullPointerException is thrown on calling the toJson method for specific Excel using DsExcel API.(DOCXLS-4829)
* NullPointerException is thrown on copying a range from a specific Excel file to a new workbook using DsExcel API.(DOCXLS-4838)
* The result in exported Excel file is incorrect after using IWorksheet.Outline.ShowLevels() method.(DOCXLS-4871)
* The cell style is incorrect in the exported PDF file.(DOCXLS-4876)
## 4.2.1
## Added
* Support for exporting charts to Image/HTML.
## Enhanced
* Enhancement of Range.AutoFit(bool considerMergedCell).
* Add dynamic array and workbook URI link to formula parser.
## Fixed                              
* ITickLabels.NumberFormat is not applied when saving a workbook to PDF/Image/HTML file.(DOCXLS-4143)
* When IAxis.MajorUnit is set, tick labels are duplicated in exported PDF/Image/HTML file.(DOCXLS-4144) 
* When the chart series is added, orders of series and legends are incorrect in exported xlsx file.(DOCXLS-4165) 
* Values of category axis are incorrect in exported PDF/Image/HTML file.(DOCXLS-4274) 
* Exception is thrown on saving PDF/Image/HTML file, if ISeries.Formula is set.(DOCXLS-4275)
* The exported file is incorrect after deleting columns by DsExcel.(DOCXLS-4606)
* Exception is thrown on using Formula parser if syntax contains '.'.(DOCXLS-4642)
* Exception is thrown on opening an Excel file.(DOCXLS-4643) 
* Cell's numberformat is different from the original JSON file in exported JSON.(DOCXLS-4653) 
* The exported CSV file is incorrect when Range.Value contains '\n'.(DOCXLS-4655)
* Exception is thrown on opening an Excel file(exported by an unknown program).(DOCXLS-4662)
* The font style is not correct in exported PDF file.(DOCXLS-4670)
* Exception should be thrown when setting Worksheet.Name which contains invalid characters.(DOCXLS-4676)
* The text in the textbox is not exported to PDF correctly.(DOCXLS-4678)
* The result of row height is incorrect in Range.RowHeight using DsExcel.(DOCXLS-4699)
* The exported JSON file by DsExcel is too large.(DOCXLS-4700)
* Exception is thrown on adding table.(DOCXLS-4702)
* Fallback font is not used after FontFolderPath is set when exporting to PDF file on Linux.(DOCXLS-4713)
* The result in exported Excel file is not correct after delete columns/rows.(DOCXLS-4728)
## 4.2.0
## Added
* Dynamic Array Formulas along with the new functions: FILTER/RANDARRAY/SEQUENCE/SINGLE/SORT/SORTBY/UNIQUE.
* Support for new Calc Engine functions: WEBSERVICE/FILTERXML/ASC/DBCS/JIS/XLOOKUP/XMATCH.
* External workbook links from the web.
* Document properties for workbook.
* Retrieve the row and column grouping information.
* Copy hidden rows to new range.
* Control the size of the exported json file.
* Support for margin settings for text in a shape.
* Expand or Collapse grouped items in a Pivot Table.
* More features for SpreadJS integration: RowCount or ColumnCount, get URL of a picture, Pivot Table for json I/O, etc.
* Support for exporting charts to PDF.
## Fixed
* NullReference exception is thrown while adding a row in a table.(DOCXLS-2504)
* When MarkerStyle.None is set for ISeries.MarkerStyle, the chart is not exported correctly in Excel.(DOCXLS-4071)
* When a JSON file is opened and exported to another JSON file, the picture floating object is lost.(DOCXLS-4098)
* Exported xlsx file does not meet the OpenXml standard if it contains a picture shape.(DOCXLS-4213)
* Exception is thrown on saving an Excel file.(DOCXLS-4215)
* When the pagebreak of template properties is set as true, page breaks are added before the field.(DOCXLS-4307)
* Dynamic array formula does not generate the correct value.(DOCXLS-4433)
* Exception is thrown on opening a JSON file.(DOCXLS-4599)
* Exception is thrown on calling the 'ToJson' method.(DOCXLS-4611)
## 4.1.4
## Fixed
* The exported PDF is different from SpreadJS designer.(DOCXLS-4152)
* Throws an exception on saving a PDF file.(DOCXLS-4407)
* Program stucks on using the 'copy' method of IRange.(DOCXLS-4414)
* 'CellInfo.GetRangeBoundary' method generates incorrect result.(DOCXLS-4424)
* Zoomfactor info is lost in the exported JSON.(DOCXLS-4427)
* The page footer in exported PDF is incorrect with the '&D' format.(DOCXLS-4432)
* The number format of PivotTable's data field is lost in the exported Excel file.(DOCXLS-4445)
* Throws an exception on using the 'IWorksheet.FromJson' method.(DOCXLS-4448)
* Treemap and Sunburst charts are not created by referring the data from another worksheet.(DOCXLS-4522)
* Drop-down list is lost in the exported JSON file.(DOCXLS-4551)
## 4.1.3
## Fixed
* Throw NumberFormatException when opening the Excel file.(DOCXLS-4146)
* Method 'toJson' fails while using Cube Formulas and OLAP Tools.(DOCXLS-4162)
* When the connector which is set as arrow head style is exported to JSON file, it is not displayed correctly on importing that JSON file in SpreadJS.(DOCXLS-4174)
* PivotTable's data field NumberFormat lost after exporting to Excel file.(DOCXLS-4214)
* Exception on opening an Excel file.(DOCXLS-4266)
* SetValue method throws exception.(DOCXLS-4273)
* An exception is thrown when opening an Excel file.(DOCXLS-4277)
* When the existing combination chart is copied in an Excel, saved xlsx file is corrupted.(DOCXLS-4287)
* Formula result calculated by DsExcel is different from Excel.(DOCXLS-4317)
* The row tags were lost in exported JSON.(DOCXLS-4321)
* The hyperlink address is incorrect in exported JSON.(DOCXLS-4322)
* Formula displays #value in DsExcel.(DOCXLS-4360)
* Opening a certain Excel file throws an Illegalargumentexception.(DOCXLS-4361)
* After copying the sheet with 'fromjson', the drop-down is lost in exported JSON.(DOCXLS-4367)
* When an xlsx file containing shape with link of a relative path is loaded and saved, saved file is corrupted.(DOCXLS-4370)
* The drop-down menu is incorrect in exported JSON.(DOCXLS-4371)
* When range is copied and Workbook.ToJson() is executed, an exception is thrown.(DOCXLS-4374)
* The cell padding style is lost in exported JSON.(DOCXLS-4391)
* The textDecoration field is missing in exported JSON when compare with the original JSON.(DOCXLS-4393)
* The JSON with cell background graph cannot be imported on loading from 'fromjson' method.(DOCXLS-4394)
## Enhanced
* Reduce the size of exported json file.(DOCXLS-4281)
* Optimize the performance of import and export json file.(DOCXLS-4372)
## 4.1.2
## Fixed
* Performance Issue in reading formula from Excel file using DsExcel API with compare to Spread.(DOCXLS-3974)
* Range Value not getting after loading file in DsExcel.(DOCXLS-4128)
* Performance issue in accessing the cell when using Slicer in the Excel sheet.(DOCXLS-4187)
* FormulaResult is not correct if the Evaluate returns null.(DOCXLS-4190)
* Saved Workbook corrupted after copying sheet with invalid cell format from existing workbook.(DOCXLS-4221)
* DsExcel for java takes too long to import ssjson.(DOCXLS-4158)
* ToJSON throw Index out of bound exception.(DOCXLS-4161)
* [I/O Lossless]Support SpreadJS table expandBoundRows.(DOCXLS-4191)
* FromJSON throw error.(DOCXLS-4212)
* The font of the exported Excel file is bold.(DOCXLS-4054)
* Setting the chart axis not to have gridlines is not applied.(DOCXLS-4100)
* Formula result is different from SpreadJS.(DOCXLS-4135)
* Program would be stuck while executing exporting pdf.(DOCXLS-4141)
* StackOverflowError occurs when importing ssjson with fromJson.(DOCXLS-4170)
* Result of exported xlsx file contains ENCODEURL function is not correct.(DOCXLS-4171)
* The value obtained by the gettext method is incorrect.(DOCXLS-4178)
* ToJSON is not including the pivotTable info.(DOCXLS-4229)
* FromJSON throws error.(DOCXLS-4235)
* After deleting a column, the getlastcolumn method does not get the correct result.(DOCXLS-4246)
* When xlsx file saved with DsExcel is loaded and saved again, number becomes incorrect.(DOCXLS-4260)
## 4.1.1
## Fixed
* Shape Color changed after saving Excel file using DsExcel API.(DOCXLS-3918)
* Invalid formula error when opening particular Excel using DsExcel API.(DOCXLS-3919)
* Behaviors of pivot table changed after saving using DsExcel API.(DOCXLS-3973)
* INDEX function value in returning "Ref" when opening excel using DsExcel API with DoNotRecalculateAfterOpened= false.(DOCXLS-3987)
* Collapsed grouping shows "-" sign when opening exported excel from DsExcel, in SpreadJs.(DOCXLS-4039)
* Value is returning "Ref" for the output fields after calling Calculate method of DsExcel.(DOCXLS-4099)
* The cell with ";;;" NumberFormat shouldn't be shown in pdf.(DOCXLS-3939)
* Formula Cells referencing Date Cells with Empty or Zero as a value are erasing the numberFormat property value after assigning a value.(DOCXLS-3990)
* The ComboBox is lost after Dsexcel loading given ssjson file.(DOCXLS-4008)
* Exception was thrown when loaing the ssjson file.(DOCXLS-4027)
* Gradient fill is displayed as solid fill in the exported HTML.(DOCXLS-3896)
* The cell value would be lost if the column's width is 1px when exporting HTML.(DOCXLS-3906)
* The drop down items is not correct in exported ssjson file.(DOCXLS-3947)
* The sheet would be lost when exporting to ssjson file.(DOCXLS-3979)
* Could not delete the custom style of DsExcel.(DOCXLS-3980)
* NullPointerException would be thrown when saving to excel file.(DOCXLS-3982)
* When an xlsx file that contains shapes is exported to pdf/image/HTML, shapes are not output correctly.(DOCXLS-3985)
* Exporting is not processed correctly when the workbook containing shapes with the same name is exported to HTML.(DOCXLS-3982)
* NullPointerException would be thrown when loaing ssjson file which contains hyperlink.(DOCXLS-3993)
* Hyperlink info would be lost in exported ssjson file.(DOCXLS-3995)
* Some exceptions would be thrown when exporting to json.(DOCXLS-4024)
* Drop down style is disorder in exported ssjson by DsExcel.(DOCXLS-4036)
* Prompt issues when open exported xlsx file by MSExcel(DOCXLS-4049).
* Infinite loop when 'PrintManager.Paginate' called with specific parameters.(DOCXLS-4056)
## 4.1.0
## Added
* Parse formula string into a syntax tree.
* Ignore Formulas when saving Excel files.
* Support open action script on PdfSaveOptions.
* New overload method to load JSON.
* More Features for SpreadJS Integration: RangeTemplate cell type, get/set custom object as cell value.
* New ToJson and FromJSON methods to Workbook elements.

## Enhanced
* Performance improvements in Excel Template processing.
* Improved Calculation Engine's performance when setting values.

## Fixed
* Performance issue when updating data in Excel using DsExcel API.(DOCXLS-3912)
* When the worksheet contains a shape, the row height cannot be changed.(DOCXLS-3909)
* The formula =TEXT("12345","[dbnum2]") not work.(DOCXLS-3900)
* The font is bold after exporting to PDF.(DOCXLS-3895)
* ROUNDUP result is different with Excel.(DOCXLS-3870)
* DsExcel formula result is different with Excel.(DOCXLS-3866)
* COUNTIF result is different with Excel.(DOCXLS-3860)
* After saving to Excel, the hidden rows are displayed.(DOCXLS-3848)
* Conditional Format is not coming in PDF/HTML.(DOCXLS-3846)
* The position of radio button list is wrong in the exported PDF.(DOCXLS-3815)
* After setting the cell style in two ways, the results are different.(DOCXLS-3787)
* Using IRange.HasFormula and IRange.Value together, degrades performance.(DOCXLS-3164)
* Mixed order of Set and Get cell values degrades performance.(DOCXLS-3163)
## 4.0.5
## Fixed
* When workbook is exported to PDF on AWS Lambda, TypeInitializationException will be thrown.(DOCXLS-3887)
* Structured references in table cells not calculating after range copy/paste between worksheets.(DOCXLS-3806)
* Incorrect value when using COUNT function with comma at last in DsExcel.(DOCXLS-3841)
* NullPointerException when Deleting the Sheet from workbook using DsExcel API.(DOCXLS-3872)
* Workbook corrupted after saving using DsExcel API.(DOCXLS-3891)
* DsExcel fails to create json from the excel file.(DOCXLS-3913)
## 4.0.4
## Fixed
* Generating SSJSON takes longer time.(DOCXLS-3598)
* Insert Range shows Invalid Operation Exception in certain cases.(DOCXLS-3638)
* ArrayIndexOutOfBoundsException while opening the file.(DOCXLS-3712)
* File takes 17min to open unless shapes disabled.(DOCXLS-3725)
* Taking longer time to get value from a cell.(DOCXLS-3736)
* Getting cell value take more time after copying it to another sheet.(DOCXLS-3738)
* Setting formula fails with exception.(DOCXLS-3742)
* Copying range fails with exception.(DOCXLS-3743)
* Getting value from source after copying takes more than 5 minutes.(DOCXLS-3744)
* Excel with pivot table corrupted after saving using DsExcel API.(DOCXLS-3747)
* NullPointerException when saving copy.(DOCXLS-3750)
* Excel corrupted after hiding columns using DsExcel API.(DOCXLS-3751)
* IndexOutOfBoundsException when saving workbook.(DOCXLS-3752)
* Cannot read formula in copied file.(DOCXLS-3765)
* Broken file produced during copy.(DOCXLS-3766)
* Removing formula in copied workbook fails with exception.(DOCXLS-3767)
* ProtectionOptions setting lost after saving to json using DsExcel.(DOCXLS-3772)
* Cached value lost in sourceSheet after copying range, Different behavior when copying range and sheet.(DOCXLS-3773)
* Calculation hangs in workbook copied using ssjson.(DOCXLS-3805)
* Structured references in table cells not calculating after range copy/paste between worksheets.(DOCXLS-3806)
* ArgumentOutOfRangeException occurs when calculating excel file with External Links.(DOCXLS-3471)
* Unexpected border in PivotTable in the exported Pdf.(DOCXLS-3770)
* Exception occurs after calling IPivotTable.Update().(DOCXLS-3797)
* Problem exporting PDF with text orientation.(DOCXLS-3455)
* IllegalArgumentException on opening attached Excel file.(DOCXLS-3472)
* Copy result is not correct if copy entire single row to multi rows and target rows contains entire merge areas.(DOCXLS-3538)
* Java API column setHidden() method is not working properly.(DOCXLS-3563)
* After the excel file toJson, data validation is lost.(DOCXLS-3649)
* After DsExcel imports and exports ssjson，the chart type has changed.(DOCXLS-3650)
* Invalid values after calculation with Release 4.0.2.(DOCXLS-3669)
* IRange.Find won't work while the search string end with "~".(DOCXLS-3708)
* Sheet password does not work after saving by DsExcel.(DOCXLS-3715)
* Exception of type 'GrapeCity.Documents.Excel.InvalidFormulaException' was thrown when Cutting a range in sheet to another workbook.(DOCXLS-3720)
* Getting the background color of entire column is wrong.(DOCXLS-3740)
* When pareto chart and funnel chart are exported to a PDF, NullReferenceException occurs.(DOCXLS-3781)
* Cells Text Orientation is not deserialised to Json.(DOCXLS-3816)
* Incorrect JSON gets exported for grouped Columns.(DOCXLS-3821)
## 4.0.3
## Fixed
* Exception occurs when save as HTML if the background image's name is null.(DOCXLS-3578)
* The line spacing becomes too big in the saved Pdf document.(DOCXLS-3524)
* When IRange.Font is getten in multi-thread processing, System.InvalidOperationException will occur.(DOCXLS-3567)
* Chart plot area in the exported Pdf is not the same as in Excel.(DOCXLS-2375)
* Some lines disappear in the exported Html.(DOCXLS-3374)
* Opening workbook fails with "Invalid argument: Number" exception.(DOCXLS-3503)
* Double underline is not exported to Pdf correctly.(DOCXLS-3507)
* Background image should not stretch in exported Html with settings "ImageLayout.None".(DOCXLS-3535)
* ClassCastException occurs when process template.(DOCXLS-3539)
* Add a picture to the sheet, and the text will not display completely in the exported Pdf.(DOCXLS-3541)
* After opening an Excel file and converting it to a json file, the word wrapping is lost.(DOCXLS-3586)
* Exception on getting formula after copying a worksheet.(DOCXLS-3587)
* Some rows are not auto-fitted in the exported Pdf.(DOCXLS-3588)
* The margins of rich text cell are not correct in the exported Pdf.(DOCXLS-3589)
* Opening file throws exception.(DOCXLS-3597)
* When a PDF form field without font name is exported to a PDF, this PDF shows an error message.(DOCXLS-3599)
* The font of cells with conditional format has changed in PDF.(DOCXLS-3602)
* ArrayIndexOutOfBoundsException in reading attached xlsx files.(DOCXLS-3628)
* When a shape is copied, horizontal alignment of text on copied shape is cleared.(DOCXLS-3630)
* Garbled text will show in some enhanced text fields.(DOCXLS-3648)
* java.lang.IllegalStateException occurs when open a xlsx file which is exported by POI.(DOCXLS-3569, DOCXLS-3572, DOCXLS-3575)
* javax.xml.stream.XMLStreamException occurs when open a xlsx file which is exported by POI.(DOCXLS-3570)
* NullPointerException occurs when open a xlsx file which is exported by POI.(DOCXLS-3571)
* java.lang.ArrayIndexOutOfBoundsException occurs when open a xlsx file which is exported by POI.(DOCXLS-3577)
## 4.0.2
## Fixed
* Image position is wrong after importing ssjson then export to PDF.(DOCXLS-3431)
* Pivot Table's border is lost in the exported PDF when there are banded rows/columns.(DOCXLS-3450)
* Exception occurs when opening certain Excel file.(DOCXLS-3479)
* Freezing range changed after saving existing excel via DsExcel API.(DOCXLS-3488)
* The image can not display when opening DsExcel exported Html in QQ and outlook mailbox.(DOCXLS-3367)
* The image border is wrong in the exported HTML.(DOCXLS-3343)
* Null exception is thrown after calling calculate method.(DOCXLS-3428)
* It becomes slow if you update cell tag in custom function.(DOCXLS-3426)
* Rows can not be deleted after calling IRange.delete method.(DOCXLS-3460)
* Open and then save an Excel with auto filter, an error occurs.(DOCXLS-3468)
* IRR formula calculation result is different with Excel.(DOCXLS-3487)
* Cell B2's value is different with Excel.(DOCXLS-3510)
* Gradient and pattern fill are lost after DsExcel load ssjson.(DOCXLS-3527)
* Copy pasting with intersecting merged cells throws exception.(DOCXLS-3161)
* When DsExcel copys whole rows/columns to other one with merged area, an invalid operation exception will occur.(DOCXLS-3205)
* The line disappears in exported image.(DOCXLS-3374)
* Sheet background image can not be exported to Html.(DOCXLS-3398)
* When the cell adjacent to the cell set the C template property is entered value or formatted, generated data is incorrect.(DOCXLS-3410)
* It takes a long time to load a ssjson.(DOCXLS-3433)
* Cell value is wrong when opening Excel with DoNotCalculateAfterOpened is true.(DOCXLS-3443)
* Exception is thrown after opening and saving Excel file.(DOCXLS-3444)
* When TemplateOptions.KeepLineSize is set true, row heights in generated xlsx file are not the same as the original template file.(DOCXLS-3448)
* ArgumentOutOfRangeException occurs when DsExcel get the cell value which can't be converted to DateTime.(DOCXLS-3457)
* ArgumentOutOfRangeException occurs when the max value of DateTime is set in a cell.(DOCXLS-3470)
* Hyperlink is not generated in exported Html when the cell is set a Hyperlink function.(DOCXLS-3475)
* Broken file created when copying sheet (related to Comments).(DOCXLS-3504)
* Cell formats are lost after SpreadJS load the ssjson generated by DsExcel.(DOCXLS-3522)
* Webdings font does not take effect in the exported PDF.(DOCXLS-3536)
## 4.0.1
## Fixed
* #REF! error occurs when copying formula between workbooks.(DOCXLS-3313)
* Broken xlsx file is generated after copying sheet between workbooks.(DOCXLS-3316,DOCXLS-3329,DOCXLS-3330)
* Hyperlinks are not exported to the PDF correctly.(DOCXLS-3335)
* Exception is thrown when saving to Excel with Pivot Table.(DOCXLS-3297)
* The col tag is not correct in the saved Html.(DOCXLS-3370)
* After saving to Html, the total value in Pivot Table shows a lot of #.(DOCXLS-3344)
* Exception is thrown during from json.(DOCXLS-3361)
* ConcurrentModificationException is thrown when saving to Excel file.(DOCXLS-3362)
* NullException is thrown when opening an Excel file.(DOCXLS-3401)
* NumberFormatException is thrown when an Excel file.(DOCXLS-3403)
* DataValidation will not work in SpreadJS after loading the ssjson exported by DsExcel.(DOCXLS-3405)
* Error occurs when opening a xlsx file which is saved by DsExcel.(DOCXLS-3408)
* IllegalStateException is thrown during from ssjson.(DOCXLS-3424)
* Column width is wrong in SpreadJS after loading the ssjson exported by DsExcel.(DOCXLS-3430)
* Can not navigate to destination address when click hyperlink in exported PDF file.(DOCXLS-3265)
* NullPointerException is thrown when copying sheet between workbooks.(DOCXLS-3315)
* Workbook.SelectedSheets doesn't reset after FromJson.(DOCXLS-3356)
* Error popping up while opening an Excel file having the pivot.(DOCXLS-3360)
* NullReferenceException will occur when Workbook.ToJson() is used.(DOCXLS-3366)
* When saving as Html, the text with format ";;;" should not be export to Html.(DOCXLS-3372)
* The custom data validation can be added without the formula.(DOCXLS-3379)
## 4.0.0
## Added
* Support for new PDF Form custom input types in Excel Templates with advanced input and validation settings.
* Support for adding, modifying and deleting Pivot Charts in Excel documents.
* Support for iterative calculations in Excel documents.
* Add Barcodes on PDF, HTML or Image Export
* Support cross-workbook formula.
* Support setting default value for template cell.
* Get range address to get cell's address.
* Add page printing events to track progress of Excel to PDF conversion.
* Select multiple worksheets.
* Get special cells in a range.
* Disable auto grouping for date/times in PivotTable.
* Add more features for SpreadJS integration: cell buttons, radio and checkbox list cell type, etc.

## Fixed
* PivotTable MergeLabel's merged area is incorrect.(DOCXLS-3181)
* PivotTable.DataBodyRange throws exception.(DOCXLS-3283)
* Can't open the xlsx when the pivot source contains null values.(DOCXLS-3285)
* Failed to export an Excel with pivot table.(DOCXLS-3297)
* Program does not end and CPU utilisation is 100% when exporting to PDF.(DOCXLS-2744)
* Broken Excel file is generated when copying a sheet.(DOCXLS-3252)
* ArrayIndexOutOfBoundsException when generating JSON.(DOCXLS-3165)
* When margins are set to the same value, the rendering position is not the same in PDF.(DOCXLS-3211)
* Labels does not merge in exported xlsx file.(DOCXLS-3262)
* SUM is calculated incorrectly when using DsExcel template functionality.(DOCXLS-3267)
* PDF form is not displayed when 'printed' and 'hidden' settings are false in form field.(DOCXLS-3275)
* Rows are not repeating while using DsExcel Template.(DOCXLS-3278)
* Null exception is thrown during loading ssjson.(DOCXLS-3290)
* Null exception is thrown when calling Workbook.Calculate.(DOCXLS-3299)
* The text in exported image is wrong.(DOCXLS-3314)
* The pivot table label does not merge in exported xlsx file.(DOCXLS-3262)

## 3.2.4
## Fixed
* DataValidation does not work for added rows for Table.(DOCXLS-3158)
* PivotTable's row field is missing after PivotTable.AddDataField().(DOCXLS-3174)
* Exception occurs when exporting Pivot Table to PDF.(DOCXLS-3176)
* PivotTable.RowRange throws exception if there is no row fields.(DOCXLS-3179)
* Table Resize with one row table lost conditional format and data validation.(DOCXLS-3194)
* The printed area does not consider hidden rows and columns.(DOCXLS-3206)
* The calculated result of Lookup formula is different with Excel.(DOCXLS-3177)
* The IRange.Columns.AutoFit() result is wrong when wrapping text.(DOCXLS-3178)
* PivotTable column fields' vertical alignment is wrong after setting MergeLabels to true.(DOCXLS-3180)
* Import ssjson and export to PDF, the program is blocked.(DOCXLS-3216)
* InvalidFormulaException is thrown when importing ssjson.(DOCXLS-3242)
* IRange.AutoFit() does not consider ConditionalFormatting's icon size.(DOCXLS-3117)
* It takes over 40 seconds to export to json.(DOCXLS-3138)
* The calculated result is wrong when array formula and custom function work together.(DOCXLS-3141)
* Richtext is lost in the exported image.(DOCXLS-3142)
* Null reference exception is thrown while exporting to json.(DOCXLS-3166)
* Null reference exception is thrown while opening an Excel with chart.(DOCXLS-3167)
* CloneNotSupportedException when copying sheet with Top/Bottom % conditional formatting rule.(DOCXLS-3169)
* Column width is not autofit correctly when Autofit() method is applied to a cell with full-width characters entered.(DOCXLS-3172)
* Grouped rows displayed in UI even after calling showLevels(1,0) after exporting to ssjson.(DOCXLS-3173)
* PivotTable's border style is wrong in the exported PDF.(DOCXLS-3175)
* The calculated style of conditional formats is wrong in the exported PDF.(DOCXLS-3234)
* NullPointerException in ArrayList occurs when copying a single cell after opening Excel File.(DOCXLS-3235)
* The Pivot Table label's resource should be "值" instead of "Values".(DOCXLS-3236)
* When workbook is exported to PDF, value axis' maximum scale in radar chart are changed.(DOCXLS-3239)
* The quote prefix is lost in the exported ssjson.(DOCXLS-3240)
* PdfSaveOptions is not applied to the output PDF when IWorksheet.Save() is used.(DOCXLS-3246)
* Spans are lost after fromJSON.(DOCXLS-3250)

## 3.2.3
## Fixed
* The generated Excel report is wrong with DsExcel teplates.(DOCXLS-2999)
* Performance issue when setting value to single cell.(DOCXLS-3029)
* The calculated result of SUMIFS formula is wrong.(DOCXLS-3033)
* Data validations are lost after expanding table ranges.(DOCXLS-3075)
* The exported xlsx can not be opened if there is an empty column in Pivot Table source.(DOCXLS-3127)
* Open a xlsx and save it back, the saved xlsx can not be opened.(DOCXLS-3131)
* The 'General' number formatter is lost in the exported ssjson.(DOCXLS-3132)
* Pagination is different between PrintManager and Workbook.Save().(DOCXLS-2801)
* Error occurs when importing JSON with some null values.(DOCXLS-3046)
* The cell text color in Pivot Talbe is wrong in the exported PDF.(DOCXLS-3130)
* The cell font size is lost after loading ssjson with Conditional Formattings.(DOCXLS-3046)
* Error occurs when loading a ssjson which contains cell value of empty array.(DOCXLS-3115)
* Cell's display color is wrong when a color number formatter is applied.(DOCXLS-3114)
* The calculated result should be #DIV/0! like Excel, but #NUM! now in DsExcel.(DOCXLS-3012)
* Out of memory occurs during processing templates.(DOCXLS-3031)
* DsExcel fails to open workbook with a checkbox wherein Excel is created using Aspose(DOCXLS-3019)
* It seems that rounding of a number does not take effect in the exported PDF.(DOCXLS-3095)
* Invalid range error occurs when setting print area.(DOCXLS-3136)
* Excel takes a long time to open a large xlsx generated by DsExcel.(DOCXLS-3128)

## 3.2.2
## Fixed
* Plot area's size and location in the exported PDF are not same as Excel.(DOCXLS-2673)
* Date/Time fields are always groupped in Pivot Table.(DOCXLS-2790)
* DataValidation does not work for added rows at 0 index for Table.(DOCXLS-2868)
* InvalidFormulaException is thrown when importing some Excel file.(DOCXLS-2889)
* Exception is thrown after processing templates and then to json.(DOCXLS-2961)
* Processing templates and exporting to json is very slow.(DOCXLS-2965)
* Incorrect settings for Protected Sheet for exported JSON from DsExcel.(DOCXLS-2977)
* Data Validation with any type in SpreadJS can't be read via DsExcel.(DOCXLS-2982)
* Error occurs when importing JSON with pictures of gif.(DOCXLS-2992)
* FM=overwrite does not work correctly when G=R in Templates.(DOCXLS-2871)
* Column width in SpreadJS is changed after DsExcel JSON I/O.(DOCXLS-2873)
* The cell with HYPERLINK formula does not display as hyperlink in the exported Excel.(DOCXLS-2914)
* Error occurs when DsExcel imports its own exported JSON .(DOCXLS-2967)
* Borders are lost in the exported PDF.(DOCXLS-2994)
* The used range becomes very large after from JSON.(DOCXLS-299)
* There are some weird lines in the exported PDF.(DOCXLS-2795)
* Chart type is changed in the exported Excel.(DOCXLS-2870)
* FM=overwrite does not work in cross-table.(DOCXLS-2872)
* The calculated result of formula "=2^10000" should be CalcError.Num.(DOCXLS-2899)
* The text of ITextRun is wrong.(DOCXLS-2902)
* The formatted text is not same as Excel with formatter "0.00E+00".(DOCXLS-2892)
* IllegalArgumentException is thrown during to JSON.(DOCXLS-2874)
* Error occurs during opening xlsx which is generated by DsExcel.(DOCXLS-2983)
* A lot of empty rows are generated after prcossing templates.(DOCXLS-2985)
* Table filter changes after saving to Excel.(DOCXLS-2962)
* Charts are changed after inserting a picture.(DOCXLS-2980)

## 3.2.1
## Fixed
* Templates markers are not cleared after processing template.(DOCXLS-2675)
* Collapsed groups are not exported correctly to ssjson.(DOCXLS-2762)
* Chart type is changed after loading a XYScatter chart from Excel.(DOCXLS-2643)
* Table Resize method does not carry over the Formula, Style and Conditional Formatting.(DOCXLS-2665)
* The exported PDF is wrong after setting best-fit-columns to true.(DOCXLS-2700)
* The exported Excel can not be opened correctly after setting IPivotTable.MergeLabels to true.(DOCXLS-2739)
* Cell text orientation is lost after importing from ssjson.(DOCXLS-2764)
* Exception is thrown when there are some special characters and exporting to PDF.(DOCXLS-2769)
* Program does not end and CPU utilisation is 100% after calling auto-fit on entire column.(DOCXLS-2785)
* Excpetion is thrown after importing a ssjson and then exporting to PDF.(DOCXLS-2789)
* The result of conditional format is wrong in the exported PDF.(DOCXLS-2685)
* The saved password-protected Excel file can not be opened when there is more than 15M data.(DOCXLS-2690)
* The height of cell(with wrapped text) is changed after importing ssjson and then saving to xlsx.(DOCXLS-2735)
* RowCount and ColumnCount in exported ssjson is incorrect after deleting some rows/columns.(DOCXLS-2741)
* Text can not be pasred correctly when current culture is en-GB.(DOCXLS-2745)
* The calculated result is "∞" in DsExcel, which is #DIV/0! in Excel.(DOCXLS-2747)
* Exception is thrown when opening an Excel with gradient unscaled fill.(DOCXLS-2761)
* The accounting format is rendered wrong in the exported PDF.(DOCXLS-2765)
* Exception is thrown when importing ssjson with an invalid "cap" value.(DOCXLS-2786)
* Range.Text is incorrect when number format is #,##0.00,,.(DOCXLS-2788)
* IRange.Clear() doesn't clear tags.(DOCXLS-2794)
* The calculated results of “=0^-x” and “=0^0” are not same as Excel.(DOCXLS-2855)

## 3.2.0
## Added
* Support for generating PDF Form from Excel Templates.
* Support using sparklines and tables in Excel Templates.
* Support defining Fixed layout for Excel report and fill data in specific range.
* Support exporting workbook/worksheet/range to HTML.
* Support Digital Signatures API: add and sign signature lines, add and sign non-visible signatures, verify signatures, etc.
* Pivot Table enhancements: create multiple files from one Pivot Field, Defer updating, Sorting, Field layout settings, etc.
* Support adjustment of shape z-order.
* Support image quality when exporting to PDF.
* Support Picture Transparency when adding Images to Excel.
* Support more SpreadJS features: show or hide horizontal and vertical grid lines, freeze trailing rows/columns, etc.

## Fixed
* Object reference error on converting the Workbook to JSON with DataValidation definitions.(DOCXLS-2659)
* ToJSON method throws error when cell has formatter on Linux.(DOCXLS-2679)
* FromJSON method takes long time to import ssjson file using DsExcel.(DOCXLS-2656)
* Some content displays incompletely in exported PDF.(DOCXLS-2737)
* Program does not end and CPU utilisation is 100% when exporting to PDF.(DOCXLS-2744)
* There is messy code in exported image generated by sheet toImage() method on Linux OS.(DOCXLS-2478)
* Program does not end and CPU utilisation is 100% when exporting to PDF.(DOCXLS-2744)
* There is messy code in exported image generated by sheet toImage() method on Linux OS.(DOCXLS-2478)

## 3.1.5
## Fixed
* Column width is not correct when workbook is exported to JSON.(DOCXLS-2497)
* NullReference exception is thrown while adding row in Table.(DOCXLS-2504)
* Background color of some cells is incorrect when rows are deleted on a sheet with background color set in cells.(DOCXLS-2530)
* Text is not correctly separated on a specific line when loading tab-delimited text to workbook.(DOCXLS-2559)
* Culture property of the Workbook does not respect the value of the cell when opening the CSV file.(DOCXLS-2566)
* NumberFormat exception is thrown when importing ssjson that contains specific conditional formats.(DOCXLS-2575)
* BlackAndWhite setting does not work for picture fill.(DOCXLS-2593)
* Rannge of the Formula is updating incorrectly when inserting row in the sheet.(DOCXLS-2595)
* The margin values in exported Excel file are slightly different with the values in ssjson.(DOCXLS-2598)
* InvalidOperationException occurs when calling Calculate method in the excel containing COUNTIFS function.(DOCXLS-2609)
* It takes a long time to set values in a loop.(DOCXLS-2610)
* Workbook.TagJsonSerializer is null when used.(DOCXLS-2606)
* processTemplate() takes a long time to execute.(DOCXLS-2526)
## 3.1.4
### Fixed
* Exception is thrown when loading Excel file which contains huge amount of inline strings.(DOCXLS-2480)
* The data labels are not plotted correctly when exporting Radar chart to PDF.(DOCXLS-2516)
* It will take a very long time to process template when there are thousands of data records.(DOCXLS-2526)
* Exception is thrown when exporting to PDF.(DOCXLS-2541)
* The calculated result of CountIfs function is wrong.(DOCXLS-2549)
## 3.1.3
### Fixed
* Error occurs when adding new row in a table.(DOCXLS-2366)
* Cell formatted text is wrong when the number format is "@".(DOCXLS-2406)
* Error occurs after calling Workbook.calculate().(DOCXLS-2430)
* Cell font color is not correct when some kind of number formatter is applied.(DOCXLS-2432)
* Cell background in the exported pdf is black even though it is "lightgray" in json.(DOCXLS-2438)
* Error occurs when from json.(DOCXLS-2440)
* Cell rich text is rendered wrong in the exported pdf when the cell is a combo cell type.(DOCXLS-2443)
* Top alignment is changed to bottom in SpreadJS after loading json from DsExcel .(DOCXLS-2469)
* Error occurs when importing Excel file with a lot of inline strings.(DOCXLS-2473)
* The saved Excel is corrupted when creating a Pivot table with SubTotalType.None.(DOCXLS-2470)
### Changed
- Depends on the latest PDFBox library-2.0.19
## 3.1.2
### Fixed
* Exception is thrown when cutting tables.(DOCXLS-2288)
* The line style in exported pdf is not same as it is in Excel.(DOCXLS-2297)
* The default date format is not changed by culture.(DOCXLS-2318)
* Exception occurs when copying a sheet from another workbook and saving the workbook.(DOCXLS-2334)
* There are unexpected lines on top of cell in the exported pdf.(DOCXLS-2339)
* The saved Excel file is corrupted.(DOCXLS-2345)
* Some values in the exported Excel file are incorrect.(DOCXLS-2352)
* It takes a long time to open an Excel with a lot of inline strings.(DOCXLS-2353)
* The plot area's boundary is incorrect in the exported pdf.(DOCXLS-2356)
* ToJson would take a very long time.(DOCXLS-2364)
* The content is incorrect in the exported pdf.(DOCXLS-2365)
* Font files will be kept opening event after saving pdf.(DOCXLS-2372)
* RepeatSetting doesn't work if there is only one row on the last page.(DOCXLS-2376])
* Picture is not clear after exporting to PDF.(DOCXLS-2384)
* Column width changes in the exported Excel.(DOCXLS-2385)
* The calculated result is incorrect.(DOCXLS-2387)
* Exception is thrown during from json.(DOCXLS-2390)
### Changed
- General formatter is used when invalid number format was found during getting cell text, previous behavior is throwing exception.
## 3.1.1
### Fixed
* Certain cell border lines do not render correctly when exporting to PDF.(DOCXLS-713)
* The order of legend for chart in exported pdf is incorrect if legend across lines.(DOCXLS-1579)
* Issue occurs when inserting new row to an existing table.(DOCXLS-1936)
* Cell Formatter change after saveJson. (DOCXLS-2013) 
* In Excel, the horizontal break line will be ignored when set the exporting height fit to 1 page(s).(DOCXLS-2097)
* IRange.Copy() cannot copy pictures sometimes.(DOCXLS-2153)
* Save Excel containing Chart shows different output.(DOCXLS-2155)
* The color of DataSeries of Chart is incorrect in exported pdf.(DOCXLS-2160)
* Name is lost after cutting table.(DOCXLS-2169)
* Stack overflow occus after cutting table.(DOCXLS-2170)
* Data validation is lost after sheet json.(DOCXLS-2175)
* Sometimes, the print area does not work.(DOCXLS-2176)
* Error occurs when fromJSON. (DOCXLS-2182)
* BackgroundPicture don't update after reset width and height if use Tile in BackgroundImageLayout. (DOCXLS-2186)
* After export PDF, a minus number is not red when use format like "0.00_;[Red]-0.00". (DOCXLS-2188)
* Picture don't update after reset width and height afert export PDF.(DOCXLS-2208)
* Formula does not change after inserting row or binding data.(DOCXLS-2220)
* The row becomes hidden after inserting row and save to json. (DOCXLS-2222)
* Error occurs during fromJSON. (DOCXLS-2227)
* ArrayIndexOutOfBounds exception occurs when cutting table.(DOCXLS-2235)
* Picture don't update after duplicating operation. (DOCXLS-2241) 
* Copying a sheet containing a chart to another workbook corrupts the book.(DOCXLS-2247)
* Calculation result is incorrect if different reference methods are used together in SUBTOTAL function.(DOCXLS-2257) 
* After export to PDF, some text can not display completely.(DOCXLS-2258)
* Error occurs after opening Excel file saved by DsExcel.(DOCXLS-2261) 
* Error occurs during from json.(DOCXLS-2262)
* Formulas is incorrectly marked as array formulas.(DOCXLS-2266)
* Data is lost after importting an Excel. (DOCXLS-2268)
* After enable PageSetup.BestFitRows, there maybe exception when save to PDF.(DOCXLS-2291)
* After enable PageSetup.BestFitRows, the exported PDF is strange.(DOCXLS-2292)
* Exception is thrown during saving to PDF when you set gradient color fill in background picture.(DOCXLS-2293)
* Background picture's gradient fill is not correct in exported PDF file.(DOCXLS-2296)
* The cell value is incorrect in the exported PDF.(DOCXLS-2303) 
* Hidden row becomes visible after save JSON.(DOCXLS-2308)
* Chart markers are not drawn correctly in exported PDF.(DOCXLS-2309)
* Incorrect calculation for CountIF formula.(DOCXLS-2310)
* There is an exception When exportting PDF.(DOCXLS-2314) 
* ComboBox CellType throw errors when fromJSON.(DOCXLS-2315)
* BackgroundPicture's gradient fill is not right after export PDF.(DOCXLS-2316)
* Cell vertical alignment is changed after from JSON.(DOCXLS-2331)
* Error occurs during from JSON. (DOCXLS-2337) 
* The large currency number is wrong in the exported PDF.(DOCXLS-2338)
* Error occurs when importing JSON with chart.(DOCXLS-2342)
## 3.1.0
### Added
* Support for charts/images/conditional formatting in Templates.
* Support exporting formulas in Templates.
* Support global settings in Templates.
* Support converting Excel objects(chart/shape) to image formats.
* Support password protected workbook and worksheet.
* Support adding Error Bars in Chart.
* Support text angle of chart title/axis tick label/data label.
* Support alignment of Shape's TextFrame.
* Add image to specific range.
* Support Gradient Fill Type enum in Shapes.
* Support creating chart/shape/pictures with a custom name.
* Enhanced Background image support for printing to PDF.
* Get pagination info for printing to PDF.
* Support Transparent Cell Background color in PDF.
* Support worksheet JSON I/O.
* Support Outline column to display hierarchical data in saved PDF.
* Support data binding of Range, Table and Worksheet.
* Return errors from JSON Import in DsExcel.
### Fixed
* Filtered data cannot be re-displayed after JSON(made by SJS) I/O in DsExcel.(DOCXLS-1646)
* Exception occurs on loading specific ssjson.(DOCXLS-2158)
* Exception may occur on exporting certain Excel sheets with charts to PDF (DOCXLS-2094)
* Hiding fixed columns and rows causes incorrect display in Excel file.(DOCXLS-2020)
* Pagination is inconsistent with SpreadJS when the form is exported to PDF.(DOCXLS-2159)
* JSON file size is bigger when converted using DsExcel vs Online designer tool.(DOCXLS-1905)
* The Value property value of the ComboBox cell is lost after JSON I/O.(DOCXLS-2029)
* NullPointerException may occur on loading certain Excel file and saving it. (DOCXLS-2041)
* Conditional formatting is lost if the rule references another sheet.(DOCXLS-2093)
## 3.0.4
### Bugs Fixed
* Exception is thrown when importing ssjson with certain group settings.(DOCXLS-1998)
* Exception is thrown when importing ssjson with some null values.(DOCXLS-1959)
* Digital signature is lost when an Excel file is modified and saved through DsExcel.(DOCXLS-1988)
* Memory leak occurs when using DsExcel in multiple threads.(DOCXLS-1950)
* Group can not be expanded in SpreadJS after loading ssjson from DsExcel.(DOCXLS-1870)
* Rich text can not be rendered correctly in the exported pdf.(DOCXLS-1906)
* The return value of IRange.HasFomula is wong.(DOCXLS-1940)
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
* Unable to import specific xlsx file in DsExcel.(DOCXLS-1848)
* Placement.MoveAndSize does not work after exporting to PDF.(DOCXLS-1856)
* External link can not work in the exported PDF when the external file's name contains Chinese characters.(DOCXLS-1867)
* Borders are lost in SpreadJS after importing ssjson generated from DsExcel.(DOCXLS-1885)
## 3.0.2
### Bugs Fixed
* Exception is thrown when importing ssjson that contains duplicated table names.(DOCXLS-1733)
* Null reference exception is thrown when exporting to ssjson.(DOCXLS-1734)
* Index out of range exception is thrown when exporting to pdf.(DOCXLS-1784)
## 3.0.1
### Bugs Fixed
* The font color is changed in SpreadJS after importing the ssjson exported by DsExcel.(DOCXLS-1729)
* The hidden rows and columns are visible in SpreadJS after importing the ssjson exported by DsExcel.(DOCXLS-1729)
* Shape name is changed after copying worksheet.(DOCXLS-1728)
* DsExcel can not load some kind of ssjson correctly. (DOCXLS-1726)
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
* Support cell tags of SpreadJS.
* Support cell types of SpreadJS.
* Support best fit rows/columns feature of SpreadJS.
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
* Calculation results from DsExcel doesn't match the results from MSExcel.(DOCXLS-1322, DOCXLS-1323)
* Cell Style is not correct after DsExcel toJSON and then SpreadJS fromJSON.(DOCXLS-1334)
## 2.2.2
### Enhancements
* Customer can overwrite current function implementation when registering a new custom function.(DOCXLS-1146)
### Bug Fixed
* Exception is thrown when exporting to json.(DOCXLS-1120)
* Exception is thrown after adding a lot of hyperlinks then saving to Excel.(DOCXLS-1154)
* Some text can not display completely in the exported pdf file.(DOCXLS-1156)
* ArgumentOutOfRangeException is thrown on using Calculate method for a Workbook.(DOCXLS-1211)
* If calculated result is too large(20+ digits) in the cell, the value acquired by DsExcel is horizontally reversed.(DOCXLS-1127)
* There is a difference between the Excel display value and the Text value acquired by DsExcel.(DOCXLS-1128)
* When a cell referenced by a formula is moved by "Cut" method, the formula is not automatically updated.(DOCXLS-1210)
* The exported pdf and Excel file is corrupted when using DsExcel in multiple threading. (DOCXLS-1208)
## 2.2.1
### Bug fixed
* DsExcel throws exception when loading ssjon with invalid image source.(DOCXLS-1115)
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
* DsExcel can not export text with Simhei bold font.(DOCXLS-937).
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
* DsExcel ignores 'ignore_empty' parameter in TEXTJOIN formula.(DOCXLS-970)
* Large JSON file generated when using ToJson on a particular Workbook.(DOCXLS-968)
* UsedRange.Value sets improper values to the range when Formula is set to Empty.(DOCXLS-956)
## 2.1.5
### Bug Fixed
* The column width, font size and picture's left position changes after JSON I/O with DsExcel.(DOCXLS-902)
* The cell font changes after JSON I/O with DsExcel.(DOCXLS-921)
* Visible property of secondary category axis returns incorrect value.(DOCXLS-849)
* The major and minor units of value axis of stacked area chart return incorrect value.(DOCXLS-805)
* The major and minor units of value axis for 100% stacked chart return incorrect value.(DOCXLS-800)
* The data label's text returns incorrect value when Axis.DisplayUnit is set. (DOCXLS-768)
* Exception is thrown on opening json file that contains invalid formula.(DOCXLS-948)
## 2.1.4
### Enhancements
* Performance of exporting a single worksheet to PDF was improved significantly.
* fromJson method  was improved when json contained table style and multiple named styles.
* DsExcel now sets minimum value of Zoom factor to 10% when exporting spreadsheet to PDF, similar to MS Excel setting. 
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
* DsExcel is unable to load an Excel file consisting SUMPRODUCT formula.(DOCXLS-733)
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
* The support for loading and saving SpreadJS JSON files with shapes have been added.
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
* DsExcel throws exception on loading ssjson file with null values.
* Merged range in table couldn't be rendered to pdf.
* The hidden rows are still rendered to pdf after loading ssjson file.


# Other Resources
* Product Home Site: [https://developer.mescius.com/document-solutions/java-excel-api](https://developer.mescius.com/document-solutions/java-excel-api)
* Demo Site: [https://developer.mescius.com/document-solutions/java-excel-api/demos/](https://developer.mescius.com/document-solutions/java-excel-api/demos/)
* Maven Repo Address: [https://search.maven.org/artifact/com.grapecity.documents/gcexcel/](https://search.maven.org/artifact/com.grapecity.documents/gcexcel/)
