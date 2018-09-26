package com.grapecity.documents.excel.examples.excelreporting;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.HorizontalAlignment;
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
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.DataLabelPosition;
import com.grapecity.documents.excel.drawing.IChartArea;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.Placement;
import com.grapecity.documents.excel.drawing.SolidColorType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PersonalNetWorthCalculator extends ExampleBase {

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

        //-------------------------Set RowHeight & Width-----------------------------------
        worksheet.setStandardHeight(30);
        worksheet.setStandardWidth(8.43);

        worksheet.getRange("1:1").setRowHeight(278.25);
        worksheet.getRange("2:4").setRowHeight(30.25);
        worksheet.getRange("8:8").setRowHeight(55.5);
        worksheet.getRange("9:30").setRowHeight(30.25);
        worksheet.getRange("33:33").setRowHeight(55.5);
        worksheet.getRange("34:44").setRowHeight(43.5);
        worksheet.getRange("A:A").setColumnWidth(2.777);
        worksheet.getRange("B:B").setColumnWidth(32.887);
        worksheet.getRange("C:C").setColumnWidth(24.219);
        worksheet.getRange("D:D").setColumnWidth(10.109);
        worksheet.getRange("E:E").setColumnWidth(61.332);
        worksheet.getRange("F:F").setColumnWidth(2.777);


        //-------------------------Set Table Value & Formulas-------------------------------
        ITable assetsTable = worksheet.getTables().add(worksheet.getRange("B9:D30"), true);
        assetsTable.setName("Assets");
        worksheet.getRange("B8").setValue("Assets");
        worksheet.getRange("B9:D30").setValue(new Object[][]{
                {"Category", "Item", "Value"},
                {"Real Estate", "Home", 560000},
                {"Real Estate", "Other", 255000},
                {"Investments", "Retirement accounts", 98000},
                {"Investments", "Stocks", 53000},
                {"Investments", "Bonds", 25000},
                {"Investments", "Mutual funds", 33000},
                {"Investments", "CDs", 74000},
                {"Investments", "Bullion", 20000},
                {"Investments", "Trust funds", 250000},
                {"Investments", "Health savings account", 18000},
                {"Investments", "Face value of life insurance policy", 85000},
                {"Investments", "Other", 20000},
                {"Cash", "Checking accounts", 14500},
                {"Cash", "Savings accounts", 5000},
                {"Cash", "Other", 2000},
                {"Personal Property", "Cars", 55000},
                {"Personal Property", "Other vehicles", 85000},
                {"Personal Property", "Furnishings", 100000},
                {"Personal Property", "Collectibles", 50000},
                {"Personal Property", "Jewelry", 60000},
                {"Personal Property", "Other luxury goods", 40000},
        });

        ITable debtsTable = worksheet.getTables().add(worksheet.getRange("B34:C44"), true);
        debtsTable.setName("Debts");
        worksheet.getRange("B33").setValue("Debts");
        worksheet.getRange("B34:C44").setValue(new Object[][]{
                {"Category", "Value"},
                {"Mortgages", 400000},
                {"Home equity loans", 50000},
                {"Car loans", 30000},
                {"Personal loans", 0},
                {"Credit cards", 0},
                {"Student loans", 10000},
                {"Loans against investments", 20000},
                {"Life insurance loans", 5000},
                {"Other installment loans", 10000},
                {"Other debts", 50000},
        });

        worksheet.getRange("B1:C1").merge();
        worksheet.getRange("B1").setValue("Personal\r\nNet\r\nWorth");
        worksheet.getRange("B2").setFormula("=\"Total \"&TotalAssetsLabel");
        worksheet.getRange("B3").setFormula("=\"Total \"&TotalDebtsLabel");
        worksheet.getRange("B4").setFormula("=NetWorthLabel");
        worksheet.getRange("C2").setFormula("=TotalAssets");
        worksheet.getRange("C3").setFormula("=TotalDebts");
        worksheet.getRange("C4").setFormula("=NetWorth");

        worksheet.getNames().add("TotalAssets", "=SUM(Assets[Value])");
        worksheet.getNames().add("TotalDebts", "=SUM(Debts[Value])");
        worksheet.getNames().add("NetWorth", "=TotalAssets-TotalDebts");
        worksheet.getNames().add("TotalAssetsLabel", "=Sheet1!$B$8");
        worksheet.getNames().add("TotalDebtsLabel", "=Sheet1!$B$33");
        worksheet.getNames().add("NetWorthLabel", "=\"Net Worth\"");


        //---------------------------Set Table Style---------------------------
        ITableStyle assetsTableStyle = workbook.getTableStyles().add("Assets");
        workbook.setDefaultTableStyle("Assets");
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getFont().setColor(Color.FromArgb(64, 64, 64));
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().setColor(Color.FromArgb(128, 128, 128));
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Dotted);
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.None);
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.None);
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.None);
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.None);

        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.SecondRowStripe).getInterior().setColor(Color.GetWhite());
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.SecondRowStripe).setStripeSize(1);

        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.LastColumn).getFont().setBold(true);
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.LastColumn).getFont().setColor(Color.FromArgb(61, 125, 137));
        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.LastColumn).getInterior().setColor(Color.GetWhite());

        assetsTableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getInterior().setColor(Color.FromArgb(61, 125, 137));


        ITableStyle debtsTableStyle = workbook.getTableStyles().add("Debts");
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getFont().setColor(Color.FromArgb(64, 64, 64));
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().setColor(Color.FromArgb(128, 128, 128));
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Dotted);
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.None);
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.None);
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.None);
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.None);

        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.SecondRowStripe).getInterior().setColor(Color.GetWhite());
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.SecondRowStripe).setStripeSize(1);

        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.LastColumn).getFont().setBold(true);
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.LastColumn).getFont().setColor(Color.FromArgb(146, 75, 12));
        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.LastColumn).getInterior().setColor(Color.GetWhite());

        debtsTableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getInterior().setColor(Color.FromArgb(218, 113, 18));


        //----------------------------Set Named Styles-------------------------
        IStyle normalStyle = workbook.getStyles().get("Normal");
        normalStyle.getFont().setName("Century Gothic");
        normalStyle.getFont().setSize(12);
        normalStyle.getFont().setColor(Color.FromArgb(64, 64, 64));
        normalStyle.getInterior().setColor(Color.FromArgb(243, 243, 236));
        normalStyle.getInterior().setPatternColor(Color.FromArgb(243, 243, 236));
        normalStyle.setHorizontalAlignment(HorizontalAlignment.Left);
        normalStyle.setIndentLevel(1);
        normalStyle.setVerticalAlignment(VerticalAlignment.Center);
        normalStyle.setWrapText(true);

        IStyle titleStyle = workbook.getStyles().get("Title");
        titleStyle.setIncludeAlignment(true);
        titleStyle.setVerticalAlignment(VerticalAlignment.Center);
        titleStyle.setWrapText(true);
        titleStyle.getFont().setName("Century Gothic");
        titleStyle.getFont().setSize(66);
        titleStyle.getFont().setColor(Color.FromArgb(64, 64, 64));
        titleStyle.setIncludePatterns(true);
        titleStyle.getInterior().setColor(Color.FromArgb(243, 243, 236));

        IStyle heading1Style = workbook.getStyles().get("Heading 1");
        heading1Style.setIncludeAlignment(true);
        heading1Style.setHorizontalAlignment(HorizontalAlignment.Left);
        heading1Style.setIndentLevel(4);
        heading1Style.setVerticalAlignment(VerticalAlignment.Center);
        heading1Style.getFont().setName("Century Gothic");
        heading1Style.getFont().setBold(false);
        heading1Style.getFont().setSize(16);
        heading1Style.getFont().setColor(Color.FromArgb(64, 64, 64));
        heading1Style.setIncludeBorder(false);
        heading1Style.setIncludePatterns(true);
        heading1Style.getInterior().setColor(Color.FromArgb(243, 243, 236));

        IStyle heading2Style = workbook.getStyles().get("Heading 2");
        heading2Style.setIncludeNumber(true);
        heading2Style.setNumberFormat("$#,##0");
        heading2Style.setIncludeAlignment(true);
        heading2Style.setHorizontalAlignment(HorizontalAlignment.Right);
        heading2Style.setIndentLevel(2);
        heading2Style.setVerticalAlignment(VerticalAlignment.Center);
        heading2Style.getFont().setName("Century Gothic");
        heading2Style.getFont().setSize(16);
        heading2Style.getFont().setColor(Color.FromArgb(64, 64, 64));
        heading2Style.setIncludeBorder(false);
        heading2Style.setIncludePatterns(true);
        heading2Style.getInterior().setColor(Color.FromArgb(243, 243, 236));

        IStyle heading3Style = workbook.getStyles().get("Heading 3");
        heading3Style.setIncludeAlignment(true);
        heading3Style.setHorizontalAlignment(HorizontalAlignment.Left);
        heading3Style.setVerticalAlignment(VerticalAlignment.Bottom);
        heading3Style.setIncludeBorder(false);
        heading3Style.getFont().setName("Century Gothic");
        heading3Style.getFont().setBold(false);
        heading3Style.getFont().setSize(27);
        heading3Style.getFont().setColor(Color.FromArgb(64, 64, 64));
        heading3Style.setIncludePatterns(true);
        heading3Style.getInterior().setColor(Color.FromArgb(243, 243, 236));

        IStyle heading4Style = workbook.getStyles().get("Heading 4");
        heading4Style.getFont().setName("Century Gothic");
        heading4Style.getFont().setSize(16);
        heading4Style.getFont().setColor(Color.GetWhite());
        heading4Style.getFont().setBold(false);

        IStyle currencyStyle = workbook.getStyles().get("Currency");
        currencyStyle.setNumberFormat("$#,##0");
        currencyStyle.setIncludeAlignment(true);
        currencyStyle.setHorizontalAlignment(HorizontalAlignment.Right);
        currencyStyle.setIndentLevel(1);
        currencyStyle.setVerticalAlignment(VerticalAlignment.Center);
        currencyStyle.setIncludeFont(true);
        currencyStyle.getFont().setBold(true);
        currencyStyle.getFont().setName("Century Gothic");
        currencyStyle.getFont().setSize(12);


        //----------------------------------Use Style---------------------------
        assetsTable.setTableStyle(assetsTableStyle);
        debtsTable.setTableStyle(debtsTableStyle);

        worksheet.getSheetView().setDisplayGridlines(false);
        worksheet.getRange("B2:B4").setStyle(heading1Style);
        worksheet.getRange("C2:C4").setStyle(heading2Style);
        worksheet.getRange("B9:D9").setStyle(heading4Style);
        worksheet.getRange("D10:D30").setStyle(currencyStyle);
        worksheet.getRange("D10:D30").getFont().setColor(Color.FromArgb(61, 125, 137));

        worksheet.getRange("B34:C34").setStyle(heading4Style);
        worksheet.getRange("C35:C44").setStyle(currencyStyle);
        worksheet.getRange("C35:C44").getFont().setColor(Color.FromArgb(218, 113, 18));
        worksheet.getRange("B1").setStyle(titleStyle);
        worksheet.getRange("B8").setStyle(heading3Style);
        worksheet.getRange("B33").setStyle(heading3Style);

        worksheet.getRange("B3:C3").getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Hair);
        worksheet.getRange("B3:C3").getBorders().get(BordersIndex.EdgeTop).setColor(Color.FromArgb(128, 128, 128));
        worksheet.getRange("B3:C3").getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Hair);
        worksheet.getRange("B3:C3").getBorders().get(BordersIndex.EdgeBottom).setColor(Color.FromArgb(128, 128, 128));


        //--------------------------------Add Shape--------------------------------
        IShape recShape1 = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 17.81, 282.75, 20.963, 21.75);
        recShape1.getLine().getColor().setColorType(SolidColorType.None);
        recShape1.getFill().getColor().setRGB(Color.FromArgb(60, 126, 138));
        IShape recShape2 = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 17.81, 312.75, 20.963, 21.75);
        recShape2.getLine().getColor().setColorType(SolidColorType.None);
        recShape2.getFill().getColor().setRGB(Color.FromArgb(218, 118, 13));
        IShape recShape3 = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 17.81, 342.75, 20.963, 21.75);
        recShape3.getLine().getColor().setColorType(SolidColorType.None);
        recShape3.getFill().getColor().setRGB(Color.FromArgb(84, 138, 57));

        IShape pieShape = worksheet.getShapes().addChart(ChartType.Pie, 442.5, 26.25, 346, 350.25);
        pieShape.getChart().setHasLegend(false);
        pieShape.getChart().setHasTitle(false);
        pieShape.getChart().getChartGroups().get(0).setFirstSliceAngle(180);
        pieShape.setPlacement(Placement.Move);

        IChartArea chartArea = pieShape.getChart().getChartArea();
        chartArea.getFormat().getFill().setTransparency(1);
        chartArea.getFormat().getLine().setTransparency(1);

        ISeries chartSeries = pieShape.getChart().getSeriesCollection().newSeries();
        chartSeries.setFormula("=SERIES('Sheet1'!$B$2:$B$4,,'Sheet1'!$C$2:$C$4,1)");

        chartSeries.setHasDataLabels(true);
        chartSeries.getDataLabels().getFont().setName("Century Gothic");
        chartSeries.getDataLabels().getFont().setSize(20);
        chartSeries.getDataLabels().getFont().setBold(true);
        chartSeries.getDataLabels().getFont().getColor().setRGB(Color.GetWhite());
        chartSeries.getDataLabels().setShowValue(false);
        chartSeries.getDataLabels().setShowPercentage(true);
        chartSeries.getDataLabels().setPosition(DataLabelPosition.Center);

        chartSeries.getPoints().get(0).getFormat().getFill().getColor().setRGB(Color.FromArgb(60, 126, 138));
        chartSeries.getPoints().get(1).getFormat().getFill().getColor().setRGB(Color.FromArgb(218, 118, 13));
        chartSeries.getPoints().get(2).getFormat().getFill().getColor().setRGB(Color.FromArgb(84, 138, 57));
        chartSeries.setExplosion(1);


    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
