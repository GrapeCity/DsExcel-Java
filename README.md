# GrapeCity Documents for Excel, Java Edition
GrapeCity introduces Documents for Excel (GcExcel) Java Edition, a high-speed, feature-rich Excel document API based on VSTO that can help developers work with spreadsheets in Java applications. The library helps to generate, convert to pdf, calculate, format, and parse spreadsheets in any application. You can work with a variety of features like importing spreadsheets, calculate data, query, generate, and export any spreadsheet, add sorting, filtering, formatting, conditional formatting and data validation, grouping, sparklines, charts, shapes, pictures, slicers, comments, hyperlinks, themes etc.  In addition, you can import existing Excel templates, add data and save the spreadsheets back. You can also use GcExcel together with Spread.Sheets, another GrapeCity Spread product that is included in GrapeCity SpreadJS. GcExcel can also import and export Excel template files on the server side. Spread.Sheets works in the browser (client side) as a viewer or editor.

With GcExcel, you can also load, edit, analyse, convert and save spreadsheets in Java applications with full support on Windows, MAC and Linux.

This repository contains source project of Examples and Showcases of GcExcel to help you learn and write your own applications. 

| Directory    | Description    |
| ------------- |-------------|
| gcexcel-1.5.3-SNAPSHOT     | Contains GcExcel Java CTP jar package and its dependency packages |
| gcexcel-1.6.0-SNAPSHOT     | Contains GcExcel Java CTP jar package and its dependency packages |
| gcexcel-2.0.0     | Contains GcExcel Java 2.0.0 jar package and its dependency packages |
| gcexcel-2.0.1     | Contains GcExcel Java 2.0.1 jar package and its dependency packages |
| gcexcel-2.1.0     | Contains GcExcel Java 2.1.0 jar package and its dependency packages |
| Examples.Library     | A collection of Java examples that help you learn and explore the API features |
| SpringBootDemo/SpringBoot+React     | A source project that demonstrates how to use GcExcel Java with SpringBoot + React + Spread.Sheets|
| SpringBootDemo/SpringBoot+Angular2     | A source project that demonstrates how to use GcExcel with SpringBoot + Angular2 + Spread.Sheets|

# Release Notes
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
