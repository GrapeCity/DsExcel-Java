package com.grapecity.documents.excel.examples.excelreporting;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.ConditionValueTypes;
import com.grapecity.documents.excel.DataBarAxisPosition;
import com.grapecity.documents.excel.DataBarDirection;
import com.grapecity.documents.excel.DataBarFillType;
import com.grapecity.documents.excel.DataBarNegativeColorType;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.IDataBar;
import com.grapecity.documents.excel.IStyle;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITableStyle;
import com.grapecity.documents.excel.ITheme;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.TableStyleElementType;
import com.grapecity.documents.excel.Theme;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.Themes;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class BidTracker extends ExampleBase {

    @Override
    public void beforeExecute(Workbook workbook, String userAgents) {

        if (userAgents.toLowerCase().contains("macintosh")) {

            ITheme theme = new Theme("testTheme", Themes.GetOfficeTheme());
            theme.getThemeFontScheme().getMinor().get(FontLanguageIndex.Latin).setName("Trebuchet MS");
            workbook.setTheme(theme);
            IStyle style_Normal = workbook.getStyles().get("Normal");
            style_Normal.getFont().setThemeFont(ThemeFont.Minor);
        }
    }

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //***********************Set RowHeight & ColumnWidth***************
        worksheet.setStandardHeight(30);
        worksheet.getRange("1:1").setRowHeight(57.75);
        worksheet.getRange("2:9").setRowHeight(30.25);
        worksheet.getRange("A:A").setColumnWidth(2.71);
        worksheet.getRange("B:B").setColumnWidth(11.71);
        worksheet.getRange("C:C").setColumnWidth(28);
        worksheet.getRange("D:D").setColumnWidth(22.425);
        worksheet.getRange("E:E").setColumnWidth(16.71);
        worksheet.getRange("F:F").setColumnWidth(28);
        worksheet.getRange("G:H").setColumnWidth(16.71);
        worksheet.getRange("I:I").setColumnWidth(2.71);

        //**************************Set Table Value & Formulas*********************
        ITable table = worksheet.getTables().add(worksheet.getRange("B2:H9"), true);
        worksheet.getRange("B2:H9").setValue(new Object[][]{
            {"BID #", "DESCRIPTION", "DATE RECEIVED", "AMOUNT", "PERCENT COMPLETE", "DEADLINE", "DAYS LEFT"},
            {1, "Bid number 1", null, 2000, 0.5, null, null},
            {2, "Bid number 2", null, 3500, 0.25, null, null},
            {3, "Bid number 3", null, 5000, 0.3, null, null},
            {4, "Bid number 4", null, 4000, 0.2, null, null},
            {5, "Bid number 5", null, 4000, 0.75, null, null},
            {6, "Bid number 6", null, 1500, 0.45, null, null},
            {7, "Bid number 7", null, 5000, 0.65, null, null},
        });
        worksheet.getRange("B1").setValue("Bid Details");
        worksheet.getRange("D3").setFormula("=TODAY()-10");
        worksheet.getRange("D4:D5").setFormula("=TODAY()-20");
        worksheet.getRange("D6").setFormula("=TODAY()-10");
        worksheet.getRange("D7").setFormula("=TODAY()-28");
        worksheet.getRange("D8").setFormula("=TODAY()-17");
        worksheet.getRange("D9").setFormula("=TODAY()-15");
        worksheet.getRange("G3:G9").setFormula("=[@[DATE RECEIVED]]+30");
        worksheet.getRange("H3:H9").setFormula("=[@DEADLINE]-TODAY()");

        //****************************Set Table Style********************************
        ITableStyle tableStyle = workbook.getTableStyles().add("Bid Tracker");
        workbook.setDefaultTableStyle("Bid Tracker");

        //Set WholeTable element style.
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getFont().setColor(Color.FromArgb(89, 89, 89));
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().setColor(Color.FromArgb(89, 89, 89));
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Thin);

        //Set HeaderRow element style.
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getFont().setColor(Color.FromArgb(89, 89, 89));
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getInterior().setColor(Color.FromArgb(131, 95, 1));
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getInterior().setPatternColor(Color.FromArgb(254, 184, 10));


        //Set TotalRow element style.
        tableStyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getBorders().setColor(Color.GetWhite());
        tableStyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getInterior().setColor(Color.FromArgb(131, 95, 1));


        //***********************************Set Named Styles*****************************
        IStyle titleStyle = workbook.getStyles().get("Title");
        titleStyle.getFont().setName("Trebuchet MS");
        titleStyle.getFont().setSize(36);
        titleStyle.getFont().setColor(Color.FromArgb(56, 145, 167));
        titleStyle.setIncludeAlignment(true);
        titleStyle.setVerticalAlignment(VerticalAlignment.Center);

        IStyle heading1Style = workbook.getStyles().get("Heading 1");
        heading1Style.setIncludeAlignment(true);
        heading1Style.setHorizontalAlignment(HorizontalAlignment.Left);
        heading1Style.setIndentLevel(1);
        heading1Style.setVerticalAlignment(VerticalAlignment.Bottom);
        heading1Style.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.None);
        heading1Style.getFont().setSize(14);
        heading1Style.getFont().setColor(Color.GetWhite());
        heading1Style.getFont().setBold(false);
        heading1Style.setIncludePatterns(true);
        heading1Style.getInterior().setColor(Color.FromArgb(131, 95, 1));
        heading1Style.getFont().setName("Trebuchet MS");


        IStyle dateStyle = workbook.getStyles().add("Date");
        dateStyle.setIncludeNumber(true);
        dateStyle.setNumberFormat("m/d/yyyy");
        dateStyle.setIncludeAlignment(true);
        dateStyle.setHorizontalAlignment(HorizontalAlignment.Left);
        dateStyle.setIndentLevel(1);
        dateStyle.setVerticalAlignment(VerticalAlignment.Center);
        dateStyle.setIncludeFont(false);
        dateStyle.setIncludeBorder(false);
        dateStyle.setIncludePatterns(false);
        dateStyle.getFont().setName("Trebuchet MS");


        IStyle commaStyle = workbook.getStyles().get("Comma");
        commaStyle.setIncludeNumber(true);
        commaStyle.setNumberFormat("#,##0_);(#,##0)");
        commaStyle.setIncludeAlignment(true);
        commaStyle.setHorizontalAlignment(HorizontalAlignment.Left);
        commaStyle.setIndentLevel(1);
        commaStyle.setVerticalAlignment(VerticalAlignment.Center);
        commaStyle.getFont().setName("Trebuchet MS");


        IStyle normalStyle = workbook.getStyles().get("Normal");
        normalStyle.setHorizontalAlignment(HorizontalAlignment.Left);
        normalStyle.setIndentLevel(1);
        normalStyle.setVerticalAlignment(VerticalAlignment.Center);
        normalStyle.setWrapText(true);
        normalStyle.getFont().setColor(Color.FromArgb(89, 89, 89));


        IStyle currencyStyle = workbook.getStyles().get("Currency");
        currencyStyle.setNumberFormat("$#,##0.00");
        currencyStyle.setIncludeAlignment(true);
        currencyStyle.setHorizontalAlignment(HorizontalAlignment.Left);
        currencyStyle.setIndentLevel(1);
        currencyStyle.setVerticalAlignment(VerticalAlignment.Center);
        currencyStyle.getFont().setName("Trebuchet MS");


        IStyle percentStyle = workbook.getStyles().get("Percent");
        percentStyle.setIncludeAlignment(true);
        percentStyle.setHorizontalAlignment(HorizontalAlignment.Right);
        percentStyle.setVerticalAlignment(VerticalAlignment.Center);
        percentStyle.setIncludeFont(true);
        percentStyle.getFont().setName("Trebuchet MS");
        percentStyle.getFont().setSize(20);
        percentStyle.getFont().setBold(true);
        percentStyle.getFont().setColor(Color.FromArgb(89, 89, 89));
        percentStyle.getFont().setName("Trebuchet MS");


        IStyle comma0Style = workbook.getStyles().get("Comma [0]");
        comma0Style.setNumberFormat("#,##0_);(#,##0)");
        comma0Style.setIncludeAlignment(true);
        comma0Style.setHorizontalAlignment(HorizontalAlignment.Right);
        comma0Style.setIndentLevel(3);
        comma0Style.setVerticalAlignment(VerticalAlignment.Center);
        percentStyle.getFont().setName("Trebuchet MS");


        //************************************Add Conditional Formatting****************
        IDataBar dataBar = worksheet.getRange("F3:F9").getFormatConditions().addDatabar();
        dataBar.getMinPoint().setType(ConditionValueTypes.Number);
        dataBar.getMinPoint().setValue(1);
        dataBar.getMaxPoint().setType(ConditionValueTypes.Number);
        dataBar.getMaxPoint().setValue(0);

        dataBar.setBarFillType(DataBarFillType.Gradient);
        dataBar.getBarColor().setColor(Color.FromArgb(126, 194, 211));
        dataBar.setDirection(DataBarDirection.Context);

        dataBar.getAxisColor().setColor(Color.GetBlack());
        dataBar.setAxisPosition(DataBarAxisPosition.Automatic);

        dataBar.getNegativeBarFormat().setColorType(DataBarNegativeColorType.Color);
        dataBar.getNegativeBarFormat().getColor().setColor(Color.GetRed());
        dataBar.setShowValue(true);


        //****************************************Use NamedStyle**************************
        worksheet.getSheetView().setDisplayGridlines(false);
        table.setTableStyle(tableStyle);
        worksheet.getRange("B1").setStyle(titleStyle);
        worksheet.getRange("B1").setWrapText(false);
        worksheet.getRange("B2:H2").setStyle(heading1Style);
        worksheet.getRange("B3:B9").setStyle(commaStyle);
        worksheet.getRange("C3:C9").setStyle(normalStyle);
        worksheet.getRange("D3:D9").setStyle(dateStyle);
        worksheet.getRange("E3:E9").setStyle(currencyStyle);
        worksheet.getRange("F3:F9").setStyle(percentStyle);
        worksheet.getRange("G3:G9").setStyle(dateStyle);
        worksheet.getRange("H3:H9").setStyle(comma0Style);

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
