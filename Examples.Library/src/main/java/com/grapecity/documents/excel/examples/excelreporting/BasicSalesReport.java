package com.grapecity.documents.excel.examples.excelreporting;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.ISlicer;
import com.grapecity.documents.excel.ISlicerCache;
import com.grapecity.documents.excel.IStyle;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITheme;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Theme;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ITextRange;
import com.grapecity.documents.excel.drawing.LineDashStyle;
import com.grapecity.documents.excel.drawing.SolidColorType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class BasicSalesReport extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //create a custom theme.
        ITheme theme = new Theme("testTheme");
        theme.getThemeColorScheme().get(ThemeColor.Light1).setRGB(Color.FromArgb(255, 255, 255));
        theme.getThemeColorScheme().get(ThemeColor.Dark1).setRGB(Color.FromArgb(0, 0, 0));
        theme.getThemeColorScheme().get(ThemeColor.Light2).setRGB(Color.FromArgb(255, 255, 255));
        theme.getThemeColorScheme().get(ThemeColor.Dark2).setRGB(Color.FromArgb(0, 0, 0));
        theme.getThemeColorScheme().get(ThemeColor.Accent1).setRGB(Color.FromArgb(140, 198, 63));
        theme.getThemeColorScheme().get(ThemeColor.Accent2).setRGB(Color.FromArgb(242, 116, 45));
        theme.getThemeColorScheme().get(ThemeColor.Accent3).setRGB(Color.FromArgb(106, 159, 207));
        theme.getThemeColorScheme().get(ThemeColor.Accent4).setRGB(Color.FromArgb(242, 192, 45));
        theme.getThemeColorScheme().get(ThemeColor.Accent5).setRGB(Color.FromArgb(146, 98, 174));
        theme.getThemeColorScheme().get(ThemeColor.Accent6).setRGB(Color.FromArgb(121, 198, 199));
        theme.getThemeColorScheme().get(ThemeColor.Hyperlink).setRGB(Color.FromArgb(106, 159, 207));
        theme.getThemeColorScheme().get(ThemeColor.FollowedHyperlink).setRGB(Color.FromArgb(146, 98, 74));
        theme.getThemeFontScheme().getMajor().get(FontLanguageIndex.Latin).setName("Garamond");
        theme.getThemeFontScheme().getMinor().get(FontLanguageIndex.Latin).setName("Garamond");

        //assign the custom theme for workbook.
        workbook.setTheme(theme);


        //Change built-in custom styles.
        IStyle style_Title = workbook.getStyles().get("Title");
        style_Title.getFont().setThemeFont(ThemeFont.Major);
        style_Title.getFont().setSize(26);
        style_Title.getFont().setBold(true);
        style_Title.getFont().setThemeColor(ThemeColor.Light1);
        style_Title.setIncludeAlignment(false);
        style_Title.setIncludeFont(true);
        style_Title.setIncludeBorder(false);
        style_Title.setIncludeNumber(false);
        style_Title.setIncludePatterns(false);
        style_Title.setIncludeProtection(false);

        IStyle style_Normal = workbook.getStyles().get("Normal");
        style_Normal.setHorizontalAlignment(HorizontalAlignment.General);
        style_Normal.setVerticalAlignment(VerticalAlignment.Center);
        style_Normal.getFont().setThemeFont(ThemeFont.Minor);
        style_Normal.getFont().setSize(9);
        style_Normal.getFont().setThemeColor(ThemeColor.Dark1);
        style_Normal.setIncludeAlignment(true);
        style_Normal.setIncludeFont(true);
        style_Normal.setIncludeBorder(true);
        style_Normal.setIncludeNumber(true);
        style_Normal.setIncludePatterns(true);
        style_Normal.setIncludeProtection(true);


        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.setName("Data Input");
        //hide worksheet gridlines.
        worksheet.getSheetView().setDisplayGridlines(false);

        //RowHeightColumnWidth
        worksheet.setStandardHeight(18.75);
        worksheet.setStandardWidth(8.43);

        worksheet.getRange("1:1").setRowHeight(51.75);
        worksheet.getRange("2:2").setRowHeight(20.25);
        worksheet.getRange("3:87").setRowHeight(19);

        worksheet.getRange("A:A").setColumnWidth(2.28515625);
        worksheet.getRange("B:B").setColumnWidth(16.85546875);
        worksheet.getRange("C:C").setColumnWidth(20.5703125);
        worksheet.getRange("D:D").setColumnWidth(27.7109375);
        worksheet.getRange("E:E").setColumnWidth(17.7109375);
        worksheet.getRange("F:F").setColumnWidth(18.140625);
        worksheet.getRange("G:G").setColumnWidth(2.28515625);

        //Values
        worksheet.getRange("B1").setValue("DATA INPUT");

        //Table
        worksheet.getRange("B2:F87").setValue(new Object[][]{
            {"DATE", "PRODUCT", "CUSTOMER", "AMOUNT", "QUARTER"},
            {new GregorianCalendar(2012, 0, 1), "Product 14", "Fabrikam, Inc.", 1886, "QUARTER 1"},
            {new GregorianCalendar(2012, 0, 3), "Product 23", "Alpine Ski House", 4022, "QUARTER 1"},
            {new GregorianCalendar(2012, 0, 4), "Product 18", "Coho Winery", 8144, "QUARTER 1"},
            {new GregorianCalendar(2012, 0, 7), "Product 10", "Southridge Video", 8002, "QUARTER 1"},
            {new GregorianCalendar(2012, 0, 11), "Product 7", "Coho Winery", 6392, "QUARTER 1"},
            {new GregorianCalendar(2012, 0, 25), "Product 1", "Contoso, Ltd", 6444, "QUARTER 1"},
            {new GregorianCalendar(2012, 0, 30), "Product 27", "Southridge Video", 2772, "QUARTER 1"},
            {new GregorianCalendar(2012, 1, 4), "Product 30", "City Power & Light", 8674, "QUARTER 1"},
            {new GregorianCalendar(2012, 1, 5), "Product 16", "A. Datum Corporation", 2332, "QUARTER 1"},
            {new GregorianCalendar(2012, 1, 8), "Product 21", "Alpine Ski House", 5370, "QUARTER 1"},
            {new GregorianCalendar(2012, 1, 10), "Product 6", "City Power & Light", 1768, "QUARTER 1"},
            {new GregorianCalendar(2012, 1, 17), "Product 24", "Coho Winery", 5474, "QUARTER 1"},
            {new GregorianCalendar(2012, 1, 22), "Product 28", "Fabrikam, Inc.", 3494, "QUARTER 1"},
            {new GregorianCalendar(2012, 1, 24), "Product 22", "City Power & Light", 1484, "QUARTER 1"},
            {new GregorianCalendar(2012, 1, 29), "Product 26", "Humongous Insurance", 5454, "QUARTER 1"},
            {new GregorianCalendar(2012, 2, 1), "Product 15", "City Power & Light", 2306, "QUARTER 1"},
            {new GregorianCalendar(2012, 2, 8), "Product 9", "A. Datum Corporation", 8652, "QUARTER 1"},
            {new GregorianCalendar(2012, 2, 14), "Product 16", "Alpine Ski House", 3594, "QUARTER 1"},
            {new GregorianCalendar(2012, 2, 31), "Product 28", "City Power & Light", 9130, "QUARTER 1"},
            {new GregorianCalendar(2012, 3, 3), "Product 28", "Southridge Video", 9986, "QUARTER 2"},
            {new GregorianCalendar(2012, 3, 9), "Product 2", "Fabrikam, Inc.", 8270, "QUARTER 2"},
            {new GregorianCalendar(2012, 3, 10), "Product 30", "A. Datum Corporation", 5184, "QUARTER 2"},
            {new GregorianCalendar(2012, 3, 11), "Product 25", "Contoso, Ltd", 9426, "QUARTER 2"},
            {new GregorianCalendar(2012, 3, 11), "Product 15", "Humongous Insurance", 4012, "QUARTER 2"},
            {new GregorianCalendar(2012, 3, 15), "Product 28", "Coho Winery", 7724, "QUARTER 2"},
            {new GregorianCalendar(2012, 4, 3), "Product 21", "Northwind Traders", 2264, "QUARTER 2"},
            {new GregorianCalendar(2012, 4, 4), "Product 30", "Coho Winery", 9374, "QUARTER 2"},
            {new GregorianCalendar(2012, 4, 5), "Product 17", "Humongous Insurance", 3692, "QUARTER 2"},
            {new GregorianCalendar(2012, 4, 5), "Product 28", "Humongous Insurance", 9124, "QUARTER 2"},
            {new GregorianCalendar(2012, 4, 19), "Product 22", "Humongous Insurance", 4962, "QUARTER 2"},
            {new GregorianCalendar(2012, 4, 22), "Product 8", "A. Datum Corporation", 9166, "QUARTER 2"},
            {new GregorianCalendar(2012, 4, 31), "Product 16", "Coho Winery", 5610, "QUARTER 2"},
            {new GregorianCalendar(2012, 5, 2), "Product 8", "City Power & Light", 3322, "QUARTER 2"},
            {new GregorianCalendar(2012, 5, 2), "Product 3", "Humongous Insurance", 2592, "QUARTER 2"},
            {new GregorianCalendar(2012, 5, 4), "Product 13", "Contoso, Ltd", 4444, "QUARTER 2"},
            {new GregorianCalendar(2012, 5, 9), "Product 10", "Southridge Video", 7166, "QUARTER 2"},
            {new GregorianCalendar(2012, 5, 12), "Product 13", "Fabrikam, Inc.", 5008, "QUARTER 3"},
            {new GregorianCalendar(2012, 5, 26), "Product 2", "Contoso, Ltd", 3578, "QUARTER 3"},
            {new GregorianCalendar(2012, 5, 29), "Product 22", "Southridge Video", 1144, "QUARTER 3"},
            {new GregorianCalendar(2012, 6, 2), "Product 14", "Alpine Ski House", 3696, "QUARTER 3"},
            {new GregorianCalendar(2012, 6, 6), "Product 14", "Coho Winery", 7084, "QUARTER 3"},
            {new GregorianCalendar(2012, 6, 6), "Product 2", "City Power & Light", 4642, "QUARTER 3"},
            {new GregorianCalendar(2012, 6, 8), "Product 5", "Fabrikam, Inc.", 6078, "QUARTER 3"},
            {new GregorianCalendar(2012, 6, 10), "Product 9", "Alpine Ski House", 2394, "QUARTER 3"},
            {new GregorianCalendar(2012, 6, 15), "Product 13", "Southridge Video", 5516, "QUARTER 3"},
            {new GregorianCalendar(2012, 6, 16), "Product 8", "Fabrikam, Inc.", 1948, "QUARTER 3"},
            {new GregorianCalendar(2012, 6, 19), "Product 28", "Contoso, Ltd", 7280, "QUARTER 3"},
            {new GregorianCalendar(2012, 6, 29), "Product 26", "A. Datum Corporation", 9292, "QUARTER 3"},
            {new GregorianCalendar(2012, 7, 19), "Product 26", "Fabrikam, Inc.", 5868, "QUARTER 3"},
            {new GregorianCalendar(2012, 7, 19), "Product 6", "Northwind Traders", 4098, "QUARTER 3"},
            {new GregorianCalendar(2012, 7, 20), "Product 13", "A. Datum Corporation", 1270, "QUARTER 3"},
            {new GregorianCalendar(2012, 7, 20), "Product 23", "A. Datum Corporation", 7744, "QUARTER 3"},
            {new GregorianCalendar(2012, 7, 24), "Product 2", "Humongous Insurance", 5488, "QUARTER 3"},
            {new GregorianCalendar(2012, 7, 24), "Product 5", "Southridge Video", 6944, "QUARTER 3"},
            {new GregorianCalendar(2012, 7, 25), "Product 20", "Fabrikam, Inc.", 4454, "QUARTER 3"},
            {new GregorianCalendar(2012, 7, 27), "Product 13", "City Power & Light", 7100, "QUARTER 3"},
            {new GregorianCalendar(2012, 7, 30), "Product 8", "Humongous Insurance", 4346, "QUARTER 3"},
            {new GregorianCalendar(2012, 8, 1), "Product 25", "Alpine Ski House", 2032, "QUARTER 3"},
            {new GregorianCalendar(2012, 8, 17), "Product 20", "Contoso, Ltd", 2328, "QUARTER 3"},
            {new GregorianCalendar(2012, 8, 22), "Product 18", "Humongous Insurance", 6090, "QUARTER 3"},
            {new GregorianCalendar(2012, 8, 25), "Product 19", "Coho Winery", 8344, "QUARTER 3"},
            {new GregorianCalendar(2012, 8, 29), "Product 11", "Alpine Ski House", 5872, "QUARTER 3"},
            {new GregorianCalendar(2012, 8, 30), "Product 25", "Humongous Insurance", 1578, "QUARTER 3"},
            {new GregorianCalendar(2012, 9, 2), "Product 29", "City Power & Light", 1714, "QUARTER 4"},
            {new GregorianCalendar(2012, 9, 5), "Product 11", "Fabrikam, Inc.", 5716, "QUARTER 4"},
            {new GregorianCalendar(2012, 9, 6), "Product 21", "Coho Winery", 8244, "QUARTER 4"},
            {new GregorianCalendar(2012, 9, 14), "Product 11", "A. Datum Corporation", 8888, "QUARTER 4"},
            {new GregorianCalendar(2012, 9, 14), "Product 23", "City Power & Light", 9438, "QUARTER 4"},
            {new GregorianCalendar(2012, 10, 9), "Product 14", "Southridge Video", 6230, "QUARTER 4"},
            {new GregorianCalendar(2012, 10, 13), "Product 17", "A. Datum Corporation", 4278, "QUARTER 4"},
            {new GregorianCalendar(2012, 10, 17), "Product 6", "Alpine Ski House", 5438, "QUARTER 4"},
            {new GregorianCalendar(2012, 10, 22), "Product 1", "Coho Winery", 6728, "QUARTER 4"},
            {new GregorianCalendar(2012, 10, 25), "Product 30", "Fabrikam, Inc.", 9992, "QUARTER 4"},
            {new GregorianCalendar(2012, 10, 25), "Product 26", "Northwind Traders", 8462, "QUARTER 4"},
            {new GregorianCalendar(2012, 10, 27), "Product 11", "Contoso, Ltd", 7930, "QUARTER 4"},
            {new GregorianCalendar(2012, 10, 30), "Product 30", "Northwind Traders", 8136, "QUARTER 4"},
            {new GregorianCalendar(2012, 11, 2), "Product 13", "Humongous Insurance", 6212, "QUARTER 4"},
            {new GregorianCalendar(2012, 11, 5), "Product 3", "Contoso, Ltd", 4946, "QUARTER 4"},
            {new GregorianCalendar(2012, 11, 5), "Product 3", "Southridge Video", 8554, "QUARTER 4"},
            {new GregorianCalendar(2012, 11, 10), "Product 24", "Northwind Traders", 4508, "QUARTER 4"},
            {new GregorianCalendar(2012, 11, 21), "Product 7", "Humongous Insurance", 7300, "QUARTER 4"},
            {new GregorianCalendar(2012, 11, 24), "Product 17", "Fabrikam, Inc.", 8292, "QUARTER 4"},
            {new GregorianCalendar(2012, 11, 26), "Product 26", "Alpine Ski House", 9782, "QUARTER 4"},
            {new GregorianCalendar(2013, 0, 3), "Product 19", "Fabrikam, Inc.", 8024, "QUARTER 1"},
            {new GregorianCalendar(2013, 0, 4), "Product 22", "A. Datum Corporation", 3758, "QUARTER 1"},
        });

        ITable table_Data = worksheet.getTables().add(worksheet.getRange("B2:F87"), true);

        //set built-in table style for table.
        table_Data.setTableStyle(workbook.getTableStyles().get("TableStyleMedium2"));

        //customize table header range's style.
        table_Data.getHeaderRange().setHorizontalAlignment(HorizontalAlignment.Left);
        table_Data.getHeaderRange().setIndentLevel(1);
        table_Data.getHeaderRange().setVerticalAlignment(VerticalAlignment.Center);
        table_Data.getHeaderRange().getFont().setSize(11);

        //customize table each column's data body range's style.
        table_Data.getColumns().get(0).getDataBodyRange().setHorizontalAlignment(HorizontalAlignment.Left);
        table_Data.getColumns().get(0).getDataBodyRange().setIndentLevel(1);
        table_Data.getColumns().get(0).getDataBodyRange().setVerticalAlignment(VerticalAlignment.Center);
        table_Data.getColumns().get(1).getDataBodyRange().setHorizontalAlignment(HorizontalAlignment.Left);
        table_Data.getColumns().get(1).getDataBodyRange().setIndentLevel(1);
        table_Data.getColumns().get(1).getDataBodyRange().setVerticalAlignment(VerticalAlignment.Center);
        table_Data.getColumns().get(2).getDataBodyRange().setHorizontalAlignment(HorizontalAlignment.Left);
        table_Data.getColumns().get(2).getDataBodyRange().setIndentLevel(1);
        table_Data.getColumns().get(2).getDataBodyRange().setVerticalAlignment(VerticalAlignment.Center);
        table_Data.getColumns().get(3).getDataBodyRange().setHorizontalAlignment(HorizontalAlignment.Right);
        table_Data.getColumns().get(3).getDataBodyRange().setIndentLevel(1);
        table_Data.getColumns().get(3).getDataBodyRange().setVerticalAlignment(VerticalAlignment.Bottom);
        table_Data.getColumns().get(3).getDataBodyRange().setNumberFormat("$#,##0.00");
        table_Data.getColumns().get(4).getDataBodyRange().setHorizontalAlignment(HorizontalAlignment.Left);
        table_Data.getColumns().get(4).getDataBodyRange().setIndentLevel(1);
        table_Data.getColumns().get(4).getDataBodyRange().setVerticalAlignment(VerticalAlignment.Bottom);


        //Slicers
        //create slicer caches.
        ISlicerCache cache_customer = workbook.getSlicerCaches().add(table_Data, "CUSTOMER", "CUSTOMER");
        ISlicerCache cache_product = workbook.getSlicerCaches().add(table_Data, "PRODUCT", "PRODUCT");

        //create slicers.
        ISlicer slicer_customer = cache_customer.getSlicers().add(worksheet, "CUSTOMER", "CUSTOMER", 51.914015748031495, 551, 144, 190);
        ISlicer slicer_product = cache_product.getSlicers().add(worksheet, "PRODUCT", "PRODUCT", 51.914015748031495, 691, 144, 190);

        //assign built-in slicer style for slicers.
        slicer_customer.setStyle(workbook.getTableStyles().get("SlicerStyleDark1"));
        slicer_product.setStyle(workbook.getTableStyles().get("SlicerStyleDark1"));


        //Style
        worksheet.getRange("1:1").setStyle(workbook.getStyles().get("Title"));
        worksheet.getRange("1:1").getInterior().setThemeColor(ThemeColor.Accent1);
        worksheet.getRange("1:1").setHorizontalAlignment(HorizontalAlignment.Left);
        worksheet.getRange("1:1").setIndentLevel(1);
        worksheet.getRange("1:1").setVerticalAlignment(VerticalAlignment.Center);
        worksheet.getRange("A1").setStyle(workbook.getStyles().get("Normal"));


        //Worksheet_CustomizableReport
        IWorksheet worksheet2 = workbook.getWorksheets().add();
        worksheet2.setName("Customizable Report");
        worksheet2.getSheetView().setDisplayGridlines(false);

        //RowHeightColumnWidth
        worksheet2.setStandardHeight(16.5);
        worksheet2.setStandardWidth(8.43);
        worksheet2.getRange("1:1").setRowHeight(51.75);
        worksheet2.getRange("2:116").setRowHeight(14.25);
        worksheet2.getRange("A:A").setColumnWidth(2.28515625);
        worksheet2.getRange("B:B").setColumnWidth(23.140625);
        worksheet2.getRange("C:C").setColumnWidth(15.5703125);
        worksheet2.getRange("D:F").setColumnWidth(11.42578125);
        worksheet2.getRange("G:H").setColumnWidth(13.42578125);

        //Values
        worksheet2.getRange("B1").setValue("SALES REPORT");


        //Style
        worksheet2.getRange("1:1").setStyle(workbook.getStyles().get("Title"));
        worksheet2.getRange("1:1").getInterior().setThemeColor(ThemeColor.Accent1);
        worksheet2.getRange("1:1").setHorizontalAlignment(HorizontalAlignment.Left);
        worksheet2.getRange("1:1").setIndentLevel(1);
        worksheet2.getRange("1:1").setVerticalAlignment(VerticalAlignment.Center);
        worksheet2.getRange("A1").setStyle(workbook.getStyles().get("Normal"));

        //Shape
        //create a shape.
        IShape shape = worksheet2.getShapes().addShape(AutoShapeType.RectangularCallout, 472, 65, 300, 70);
        //config shape's line and fill.
        shape.getLine().setDashStyle(LineDashStyle.Solid);
        shape.getLine().setWeight(4);
        shape.getLine().getColor().setRGB(Color.FromArgb(89, 89, 89));
        shape.getFill().solid();
        shape.getFill().getColor().setColorType(SolidColorType.None);

        //config shape's rich text.
        ITextRange shape_p1 = shape.getTextFrame().getTextRange().getParagraphs().get(0);
        shape_p1.setText("TIP:");
        ITextRange shape_p2 = shape.getTextFrame().getTextRange().getParagraphs().add();
        shape_p2.setText("Customize this PivotTable to fit your needs! Select a cell in the PivotTable to activate the PivotTable Field List pane. Then in the task pane, drag to add, remove, or reorder the fields." +
            " For example, drag the Product field above the Customer field for a different view. To update PivotTable data, right-click in the PivotTable and then click Refresh.");

        //set first paragraph's font style.
        shape_p1.getFont().setThemeFont(ThemeFont.Major);
        shape_p1.getFont().setBold(true);
        shape_p1.getFont().setSize(12);
        shape_p1.getFont().getColor().setObjectThemeColor(ThemeColor.Dark1);
        shape_p1.getFont().getColor().setBrightness(0.25);

        //set second paragraph's font style.
        shape_p2.getFont().setThemeFont(ThemeFont.Minor);
        shape_p2.getFont().setSize(8);
        shape_p2.getFont().getColor().setObjectThemeColor(ThemeColor.Dark1);
        shape_p2.getFont().getColor().setBrightness(0.25);


        //Worksheet_DataLists
        IWorksheet worksheet3 = workbook.getWorksheets().add();
        worksheet3.setName("Data Lists");
        worksheet3.getSheetView().setDisplayGridlines(false);

        //RowHeightColumnWidth
        worksheet3.setStandardHeight(18.75);
        worksheet3.setStandardWidth(8.43);
        worksheet3.getRange("1:1").setRowHeight(51.75);
        worksheet3.getRange("2:32").setRowHeight(19);
        worksheet3.getRange("A:A, D:D").setColumnWidth(2.28515625);
        worksheet3.getRange("B:C").setColumnWidth(34.42578125);

        //Values
        worksheet3.getRange("B1").setValue("DATA LISTS");

        //Table
        worksheet3.getRange("B2:B32").setValue(new Object[][]{
            {"PRODUCTS"},
            {"Product 1"},
            {"Product 2"},
            {"Product 3"},
            {"Product 4"},
            {"Product 5"},
            {"Product 6"},
            {"Product 7"},
            {"Product 8"},
            {"Product 9"},
            {"Product 10"},
            {"Product 11"},
            {"Product 12"},
            {"Product 13"},
            {"Product 14"},
            {"Product 15"},
            {"Product 16"},
            {"Product 17"},
            {"Product 18"},
            {"Product 19"},
            {"Product 20"},
            {"Product 21"},
            {"Product 22"},
            {"Product 23"},
            {"Product 24"},
            {"Product 25"},
            {"Product 26"},
            {"Product 27"},
            {"Product 28"},
            {"Product 29"},
            {"Product 30"},
        });
        ITable table_Products = worksheet3.getTables().add(worksheet.getRange("B2:B32"), true);
        worksheet3.getRange("C2:C30").setValue(new Object[][]{
            {"CUSTOMERS"},
            {"A. Datum Corporation"},
            {"Adventure Works"},
            {"Alpine Ski House"},
            {"Blue Yonder Airlines"},
            {"City Power & Light"},
            {"Coho Vineyard"},
            {"Coho Winery"},
            {"Coho Vineyard & Winery"},
            {"Contoso, Ltd"},
            {"Contoso Pharmaceuticals"},
            {"Consolidated Messenger"},
            {"Fabrikam, Inc."},
            {"Fourth Coffee"},
            {"Graphic Design Institute"},
            {"Humongous Insurance"},
            {"Litware, Inc."},
            {"Lucerne Publishing"},
            {"Margie's Travel"},
            {"Northwind Traders"},
            {"Proseware, Inc."},
            {"School of Fine Art"},
            {"Southridge Video"},
            {"Tailspin Toys"},
            {"Trey Research"},
            {"The Phone Company"},
            {"Wide World Importers"},
            {"Wingtip Toys"},
            {"Woodgrove Bank"},
        });
        ITable table_Customers = worksheet3.getTables().add(worksheet.getRange("C2:C30"), true);

        table_Products.setTableStyle(workbook.getTableStyles().get("TableStyleMedium2"));
        table_Products.getColumns().get(0).getRange().setHorizontalAlignment(HorizontalAlignment.Left);
        table_Products.getColumns().get(0).getRange().setIndentLevel(1);
        table_Products.getColumns().get(0).getRange().setVerticalAlignment(VerticalAlignment.Center);
        table_Products.getHeaderRange().getFont().setSize(11);

        table_Customers.setTableStyle(workbook.getTableStyles().get("TableStyleMedium2"));
        table_Customers.getColumns().get(0).getRange().setHorizontalAlignment(HorizontalAlignment.Left);
        table_Customers.getColumns().get(0).getRange().setIndentLevel(1);
        table_Customers.getColumns().get(0).getRange().setVerticalAlignment(VerticalAlignment.Center);
        table_Customers.setShowTableStyleFirstColumn(true);
        table_Customers.getHeaderRange().getFont().setSize(11);

        //Style
        worksheet3.getRange("1:1").setStyle(workbook.getStyles().get("Title"));
        worksheet3.getRange("1:1").getInterior().setThemeColor(ThemeColor.Accent1);
        worksheet3.getRange("1:1").setHorizontalAlignment(HorizontalAlignment.Left);
        worksheet3.getRange("1:1").setIndentLevel(1);
        worksheet3.getRange("1:1").setVerticalAlignment(VerticalAlignment.Center);
        worksheet3.getRange("A1").setStyle(workbook.getStyles().get("Normal"));

        //Shape
        IShape shape2 = worksheet3.getShapes().addShape(AutoShapeType.RectangularCallout, 380, 65, 280, 50);
        shape2.getLine().setDashStyle(LineDashStyle.Solid);
        shape2.getLine().setWeight(4);
        shape2.getLine().getColor().setRGB(Color.FromArgb(89, 89, 89));
        shape2.getFill().solid();
        shape2.getFill().getColor().setColorType(SolidColorType.None);

        ITextRange shape2_p1 = shape2.getTextFrame().getTextRange().getParagraphs().get(0);
        shape2_p1.setText("TIP:");
        ITextRange shape2_p2 = shape2.getTextFrame().getTextRange().getParagraphs().add();
        shape2_p2.setText("To add a new product or customer, start typing below the table and it will automatically expand when you press the Enter or Tab key.");
        shape2_p1.getFont().setThemeFont(ThemeFont.Major);
        shape2_p1.getFont().setBold(true);
        shape2_p1.getFont().setSize(12);
        shape2_p1.getFont().getColor().setObjectThemeColor(ThemeColor.Dark1);
        shape2_p1.getFont().getColor().setBrightness(0.25);

        shape2_p2.getFont().setThemeFont(ThemeFont.Minor);
        shape2_p2.getFont().setSize(8);
        shape2_p2.getFont().getColor().setObjectThemeColor(ThemeColor.Dark1);
        shape2_p2.getFont().getColor().setBrightness(0.25);

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
