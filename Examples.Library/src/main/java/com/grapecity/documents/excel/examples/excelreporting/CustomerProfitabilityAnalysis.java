package com.grapecity.documents.excel.examples.excelreporting;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.ISparklineGroup;
import com.grapecity.documents.excel.IStyle;
import com.grapecity.documents.excel.ITheme;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.SparkType;
import com.grapecity.documents.excel.Theme;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.IChartTitle;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.LegendPosition;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CustomerProfitabilityAnalysis extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        Object data = new Object[][]{
            {null, "[Segment Name]", "[Segment Name]", "[Segment Name]", "Overall"},
            {"Customer Activity:", null, null, null, null},
            {"Number of active customers—Beginning of period", 5, 8, 8, null},
            {"Number of customers added", 2, 4, 4, null},
            {"Number of customers lost/terminated", -1, -2, -2, null},
            {"Number of active customers—End of period", null, null, null, null},
            {null, null, null, null, null},
            {"Profitability Analysis:", null, null, null, null},
            {"Revenue per segment", 1500000, 1800000, 2500000, null},
            {"Weighting", null, null, null, null},
            {null, null, null, null, null},
            {"Cost of sales:", null, null, null, null},
            {"Ongoing service and support costs", 1000000, 1400000, 1400000, null},
            {"Other direct customer costs", 200000, 100000, 100000, null},
            {"Total cost of sales", null, null, null, null},
            {"Gross margin", null, null, null, null},
            {"Weighting", null, null, null, null},
            {null, null, null, null, null},
            {"Other costs:", null, null, null, null},
            {"Customer acquisition", 105000, 120000, 235000, null},
            {"Customer marketing", 150000, 125000, 275000, null},
            {"Customer termination", 80000, 190000, 140000, null},
            {"Total other customer costs", null, null, null, null},
            {"Customer profit by segment", null, null, null, null},
            {"Weighting", null, null, null, null},
            {null, null, null, null, null},
            {"Summary Metrics:", "[Segment Name]", "[Segment Name]", "[Segment Name]", "Trend"},
            {"Average cost per acquired customer", null, null, null, null},
            {"Average cost per terminated customer", null, null, null, null},
            {"Average marketing cost per active customer", null, null, null, null},
            {"Average profit (loss) per customer", null, null, null, null},
        };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.setName("Customer Profitability");
        worksheet.setTabColor(Color.FromArgb(131, 172, 121));
        worksheet.getSheetView().setDisplayGridlines(false);

        //Set Value.
        worksheet.getRange("B2").setValue("[Company Name]");
        worksheet.getRange("B3").setValue("Customer Profitability Analysis");
        worksheet.getRange("B4").setValue("[Date]");
        worksheet.getRange("B6").setValue("Gray cells will be calculated for you. You do not need to enter anything into them.");
        worksheet.getRange("B7:F37").setValue(data);

        //Set formula.
        worksheet.getRange("F9:F11").setFormula("=SUM(C9:E9)");
        worksheet.getRange("C12:F12").setFormula("=SUM(C9:C11)");

        worksheet.getRange("C16:E16").setFormula("=+C15/$F$15");
        worksheet.getRange("F15:F16").setFormula("=SUM(C15:E15)");

        worksheet.getRange("F19:F20").setFormula("=SUM(C19:E19)");
        worksheet.getRange("C21:F21").setFormula("=SUM(C19:C20)");
        worksheet.getRange("C22:F22").setFormula("=+C15-C21");
        worksheet.getRange("C23:E23").setFormula("=MAX(0, MIN(1,C22/$F$22))");
        worksheet.getRange("F23").setFormula("=SUM(C23:E23)");

        worksheet.getRange("F26:F28").setFormula("=SUM(C26:E26)");
        worksheet.getRange("C29:F29").setFormula("=SUM(C26:C28)");
        worksheet.getRange("C30:E30").setFormula("=+C22-C29");
        worksheet.getRange("C31:E31").setFormula("=MAX(0,MIN(1, C30/$F$30))");
        worksheet.getRange("F30:F31").setFormula("=SUM(C30:E30)");

        worksheet.getRange("C34:E34").setFormula("=+C26/C10");
        worksheet.getRange("C35:E35").setFormula("=-C28/C11");
        worksheet.getRange("C36:E36").setFormula("=+C27/C12");
        worksheet.getRange("C37:E37").setFormula("=+C30/C12");

        //Change the range's RowHeight and ColumnWidth.
        worksheet.setStandardHeight(15);
        worksheet.setStandardWidth(9.140625);
        worksheet.getRows().get(0).setRowHeight(9.95);
        worksheet.getRows().get(1).setRowHeight(33);
        worksheet.getRows().get(2).setRowHeight(27);
        worksheet.getRows().get(3).setRowHeight(19.5);
        worksheet.getRows().get(4).setRowHeight(9);
        worksheet.getRows().get(5).setRowHeight(19.5);
        worksheet.getRows().get(6).setRowHeight(18);
        worksheet.getRows().get(12).setRowHeight(9);
        worksheet.getRows().get(16).setRowHeight(9);
        worksheet.getRows().get(23).setRowHeight(9);
        worksheet.getRows().get(31).setRowHeight(9);

        worksheet.getColumns().get(0).setColumnWidth(1.85546875);
        worksheet.getColumns().get(1).setColumnWidth(46.7109375);
        worksheet.getColumns().get(2).setColumnWidth(16.42578125);
        worksheet.getColumns().get(3).setColumnWidth(16.42578125);
        worksheet.getColumns().get(4).setColumnWidth(16.42578125);
        worksheet.getColumns().get(5).setColumnWidth(16.42578125);

        //Modify the build in name styles.
        IStyle nameStyle_Normal = workbook.getStyles().get("Normal");
        nameStyle_Normal.setVerticalAlignment(VerticalAlignment.Center);
        nameStyle_Normal.getFont().setThemeColor(ThemeColor.Dark1);
        nameStyle_Normal.getFont().setTintAndShade(0.249946592608417);
        nameStyle_Normal.getFont().setSize(10);

        IStyle nameStyle_Heading_1 = workbook.getStyles().get("Heading 1");
        nameStyle_Heading_1.setHorizontalAlignment(HorizontalAlignment.Left);
        nameStyle_Heading_1.setVerticalAlignment(VerticalAlignment.Center);
        nameStyle_Heading_1.getFont().setThemeFont(ThemeFont.Major);
        nameStyle_Heading_1.getFont().setBold(false);
        nameStyle_Heading_1.getFont().setSize(24);
        nameStyle_Heading_1.getFont().setThemeColor(ThemeColor.Dark1);
        nameStyle_Heading_1.getFont().setTintAndShade(0.249946592608417);
        nameStyle_Heading_1.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.None);
        nameStyle_Heading_1.setIncludeAlignment(true);

        IStyle nameStyle_Heading_2 = workbook.getStyles().get("Heading 2");
        nameStyle_Heading_2.setHorizontalAlignment(HorizontalAlignment.Left);
        nameStyle_Heading_2.setVerticalAlignment(VerticalAlignment.Center);
        nameStyle_Heading_2.getFont().setThemeFont(ThemeFont.Major);
        nameStyle_Heading_2.getFont().setBold(false);
        nameStyle_Heading_2.getFont().setSize(20);
        nameStyle_Heading_2.getFont().setThemeColor(ThemeColor.Dark1);
        nameStyle_Heading_2.getFont().setTintAndShade(0.249946592608417);
        nameStyle_Heading_2.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.None);
        nameStyle_Heading_2.getInterior().setThemeColor(ThemeColor.Accent3);
        nameStyle_Heading_2.getInterior().setTintAndShade(0.39994506668294322);
        nameStyle_Heading_2.setIncludeNumber(true);
        nameStyle_Heading_2.setIncludePatterns(true);

        IStyle nameStyle_Heading_3 = workbook.getStyles().get("Heading 3");
        nameStyle_Heading_3.setHorizontalAlignment(HorizontalAlignment.Left);
        nameStyle_Heading_3.setVerticalAlignment(VerticalAlignment.Center);
        nameStyle_Heading_3.getFont().setThemeFont(ThemeFont.Major);
        nameStyle_Heading_3.getFont().setBold(false);
        nameStyle_Heading_3.getFont().setSize(14);
        nameStyle_Heading_3.getFont().setThemeColor(ThemeColor.Dark1);
        nameStyle_Heading_3.getFont().setTintAndShade(0.249946592608417);
        nameStyle_Heading_3.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.None);
        nameStyle_Heading_3.setIncludeAlignment(true);
        nameStyle_Heading_3.setIncludePatterns(true);

        IStyle nameStyle_Heading_4 = workbook.getStyles().get("Heading 4");
        nameStyle_Heading_4.setHorizontalAlignment(HorizontalAlignment.Left);
        nameStyle_Heading_4.setVerticalAlignment(VerticalAlignment.Center);
        nameStyle_Heading_4.getFont().setThemeFont(ThemeFont.Major);
        nameStyle_Heading_4.getFont().setBold(true);
        nameStyle_Heading_4.getFont().setSize(10);
        nameStyle_Heading_4.getFont().setThemeColor(ThemeColor.Light1);
        nameStyle_Heading_4.getFont().setTintAndShade(-0.0499893185216834);
        nameStyle_Heading_4.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.None);
        nameStyle_Heading_4.getInterior().setThemeColor(ThemeColor.Accent3);
        nameStyle_Heading_4.getInterior().setTintAndShade(-0.249946592608417);
        nameStyle_Heading_4.setIncludeAlignment(true);
        nameStyle_Heading_4.setIncludeBorder(true);
        nameStyle_Heading_4.setIncludePatterns(true);

        //Apply the above name styles on ranges.
        worksheet.getRange("B2:F2").setStyle(workbook.getStyles().get("Heading 1"));
        worksheet.getRange("B3:F3").setStyle(workbook.getStyles().get("Heading 2"));
        worksheet.getRange("B4:F4").setStyle(workbook.getStyles().get("Heading 3"));
        worksheet.getRange("B8:F8").setStyle(workbook.getStyles().get("Heading 4"));
        worksheet.getRange("B14:F14").setStyle(workbook.getStyles().get("Heading 4"));
        worksheet.getRange("B18:F18").setStyle(workbook.getStyles().get("Heading 4"));
        worksheet.getRange("B25:F25").setStyle(workbook.getStyles().get("Heading 4"));
        worksheet.getRange("B33:F33").setStyle(workbook.getStyles().get("Heading 4"));

        //Set NumberFormat.
        worksheet.getRange("C9:F12").setNumberFormat("0_);[Red](0)");
        worksheet.getRange("C15:F15").setNumberFormat("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)");
        worksheet.getRange("C16:F16").setNumberFormat("0%");
        worksheet.getRange("C19:F22").setNumberFormat("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)");
        worksheet.getRange("C23:F23").setNumberFormat("0%");
        worksheet.getRange("C26:F30").setNumberFormat("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)");
        worksheet.getRange("C31:F31").setNumberFormat("0%");
        worksheet.getRange("C34:F37").setNumberFormat("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)");

        //Set range's font style.
        worksheet.getRange("B6").getFont().setTintAndShade(0.34998626667073579);
        worksheet.getRange("B6").getFont().setSize(8);
        worksheet.getRange("B6").getFont().setItalic(true);
        worksheet.getRange("C7:F7").getFont().setTintAndShade(0);
        worksheet.getRange("B9:F12").getFont().setTintAndShade(0);
        worksheet.getRange("B15:F16").getFont().setTintAndShade(0);
        worksheet.getRange("B19:F23").getFont().setTintAndShade(0);
        worksheet.getRange("B26:F31").getFont().setTintAndShade(0);
        worksheet.getRange("B34:F37").getFont().setTintAndShade(0);
        worksheet.getRange("C33:F33").getFont().setBold(false);

        //Set range's alignment.
        worksheet.getRange("C7:F7").setHorizontalAlignment(HorizontalAlignment.Center);
        worksheet.getRange("C33:F33").setHorizontalAlignment(HorizontalAlignment.Center);

        //Set range's border
        worksheet.getRange("B9:F12").getBorders().setLineStyle(BorderLineStyle.Thin);
        worksheet.getRange("B9:F12").getBorders().setThemeColor(ThemeColor.Accent3);
        worksheet.getRange("B9:F12").getBorders().setTintAndShade(0.39994506668294322);

        worksheet.getRange("B15:F16").getBorders().setLineStyle(BorderLineStyle.Thin);
        worksheet.getRange("B15:F16").getBorders().setThemeColor(ThemeColor.Accent3);
        worksheet.getRange("B15:F16").getBorders().setTintAndShade(0.39994506668294322);

        worksheet.getRange("B19:F23").getBorders().setLineStyle(BorderLineStyle.Thin);
        worksheet.getRange("B19:F23").getBorders().setThemeColor(ThemeColor.Accent3);
        worksheet.getRange("B19:F23").getBorders().setTintAndShade(0.39994506668294322);

        worksheet.getRange("B26:F31").getBorders().setLineStyle(BorderLineStyle.Thin);
        worksheet.getRange("B26:F31").getBorders().setThemeColor(ThemeColor.Accent3);
        worksheet.getRange("B26:F31").getBorders().setTintAndShade(0.39994506668294322);

        worksheet.getRange("B34:F37").getBorders().setLineStyle(BorderLineStyle.Thin);
        worksheet.getRange("B34:F37").getBorders().setThemeColor(ThemeColor.Accent3);
        worksheet.getRange("B34:F37").getBorders().setTintAndShade(0.39994506668294322);

        //Set range's fill.
        worksheet.getRange("F9:F12").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("F9:F12").getInterior().setTintAndShade(-0.0499893185216834);
        worksheet.getRange("C12:E12").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("C12:E12").getInterior().setTintAndShade(-0.0499893185216834);
        worksheet.getRange("F15:F16").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("F15:F16").getInterior().setTintAndShade(-0.0499893185216834);
        worksheet.getRange("C16:E16").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("C16:E16").getInterior().setTintAndShade(-0.0499893185216834);
        worksheet.getRange("F19:F23").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("F19:F23").getInterior().setTintAndShade(-0.0499893185216834);
        worksheet.getRange("C21:E23").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("C21:E23").getInterior().setTintAndShade(-0.0499893185216834);
        worksheet.getRange("F26:F31").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("F26:F31").getInterior().setTintAndShade(-0.0499893185216834);
        worksheet.getRange("C29:E31").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("C29:E31").getInterior().setTintAndShade(-0.0499893185216834);
        worksheet.getRange("C34:E37").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("C34:E37").getInterior().setTintAndShade(-0.0499893185216834);

        //create a new group of sparklines.
        ISparklineGroup sparklineGroup = worksheet.getRange("F34:F37").getSparklineGroups().add(SparkType.Line, "C34:E37");

        sparklineGroup.getSeriesColor().setThemeColor(ThemeColor.Accent3);
        sparklineGroup.getSeriesColor().setTintAndShade(-0.249977111117893);
        sparklineGroup.getPoints().getNegative().getColor().setThemeColor(ThemeColor.Accent4);
        sparklineGroup.getPoints().getMarkers().getColor().setThemeColor(ThemeColor.Accent4);
        sparklineGroup.getPoints().getMarkers().getColor().setTintAndShade(-0.249977111117893);
        sparklineGroup.getPoints().getHighpoint().getColor().setThemeColor(ThemeColor.Accent4);
        sparklineGroup.getPoints().getHighpoint().getColor().setTintAndShade(-0.249977111117893);
        sparklineGroup.getPoints().getLowpoint().getColor().setThemeColor(ThemeColor.Accent4);
        sparklineGroup.getPoints().getLowpoint().getColor().setTintAndShade(-0.249977111117893);
        sparklineGroup.getPoints().getFirstpoint().getColor().setThemeColor(ThemeColor.Accent4);
        sparklineGroup.getPoints().getFirstpoint().getColor().setTintAndShade(-0.249977111117893);
        sparklineGroup.getPoints().getLastpoint().getColor().setThemeColor(ThemeColor.Accent4);
        sparklineGroup.getPoints().getLastpoint().getColor().setTintAndShade(-0.249977111117893);
        sparklineGroup.getPoints().getNegative().setVisible(false);
        sparklineGroup.getPoints().getFirstpoint().setVisible(false);
        sparklineGroup.getPoints().getLastpoint().setVisible(false);

        //Add chart.
        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 9.75, 576.95, 590.25, 237);
        shape.setName("Chart 3");

        //Add Series.
        ISeries series1 = shape.getChart().getSeriesCollection().newSeries();
        series1.setFormula("=SERIES('Customer Profitability'!$B$34,'Customer Profitability'!$C$33:$E$33,'Customer Profitability'!$C$34:$E$34,1)");
        series1.getFormat().getFill().getColor().setObjectThemeColor(ThemeColor.Accent2);

        ISeries series2 = shape.getChart().getSeriesCollection().newSeries();
        series2.setFormula("=SERIES('Customer Profitability'!$B$35,'Customer Profitability'!$C$33:$E$33,'Customer Profitability'!$C$35:$E$35,2)");
        series2.getFormat().getFill().getColor().setObjectThemeColor(ThemeColor.Accent4);

        ISeries series3 = shape.getChart().getSeriesCollection().newSeries();
        series3.setFormula("=SERIES('Customer Profitability'!$B$36,'Customer Profitability'!$C$33:$E$33,'Customer Profitability'!$C$36:$E$36,3)");
        series3.getFormat().getFill().getColor().setObjectThemeColor(ThemeColor.Accent3);

        ISeries series4 = shape.getChart().getSeriesCollection().newSeries();
        series4.setFormula("=SERIES('Customer Profitability'!$B$37,'Customer Profitability'!$C$33:$E$33,'Customer Profitability'!$C$37:$E$37,4)");
        series4.getFormat().getFill().getColor().setObjectThemeColor(ThemeColor.Accent5);

        //Set the char group's Overlap and GapWidth.
        shape.getChart().getColumnGroups().get(0).setOverlap(0);
        shape.getChart().getColumnGroups().get(0).setGapWidth(199);

        //Set the chart's title format.
        IChartTitle chartTitle = shape.getChart().getChartTitle();
        chartTitle.setText("Summary Metrics per Customer Segment");
        chartTitle.getFont().setThemeFont(ThemeFont.Major);
        chartTitle.getFont().getColor().setObjectThemeColor(ThemeColor.Dark1);
        chartTitle.getFont().setSize(20);

        //Set the chart legend's position.
        shape.getChart().getLegend().setPosition(LegendPosition.Top);

        //Set category axis format.
        IAxis category_axis = shape.getChart().getAxes().item(AxisType.Category);
        category_axis.setHasTitle(true);
        category_axis.getAxisTitle().setText("SEGMENT");
        category_axis.getAxisTitle().getFont().setSize(9);
        category_axis.getAxisTitle().getFont().setThemeFont(ThemeFont.Minor);

        //Set value axis format.
        IAxis value_axis = shape.getChart().getAxes().item(AxisType.Value);
        value_axis.setCrossesAt(-200000);
        value_axis.setHasMinorGridlines(true);
        value_axis.getMinorGridlines().getFormat().getLine().getColor().setObjectThemeColor(ThemeColor.Dark1);
        value_axis.getMinorGridlines().getFormat().getLine().getColor().setBrightness(0.95);

        //Create customize theme.
        ITheme theme = new Theme("test");
        theme.getThemeColorScheme().get(ThemeColor.Dark1).setRGB(Color.FromArgb(0, 0, 0));
        theme.getThemeColorScheme().get(ThemeColor.Light1).setRGB(Color.FromArgb(255, 255, 255));
        theme.getThemeColorScheme().get(ThemeColor.Dark2).setRGB(Color.FromArgb(77, 70, 70));
        theme.getThemeColorScheme().get(ThemeColor.Light2).setRGB(Color.FromArgb(255, 251, 239));
        theme.getThemeColorScheme().get(ThemeColor.Accent1).setRGB(Color.FromArgb(255, 225, 132));
        theme.getThemeColorScheme().get(ThemeColor.Accent2).setRGB(Color.FromArgb(102, 173, 166));
        theme.getThemeColorScheme().get(ThemeColor.Accent3).setRGB(Color.FromArgb(131, 172, 121));
        theme.getThemeColorScheme().get(ThemeColor.Accent4).setRGB(Color.FromArgb(254, 191, 102));
        theme.getThemeColorScheme().get(ThemeColor.Accent5).setRGB(Color.FromArgb(219, 112, 87));
        theme.getThemeColorScheme().get(ThemeColor.Accent6).setRGB(Color.FromArgb(165, 115, 137));
        theme.getThemeColorScheme().get(ThemeColor.Hyperlink).setRGB(Color.FromArgb(102, 173, 166));
        theme.getThemeColorScheme().get(ThemeColor.FollowedHyperlink).setRGB(Color.FromArgb(165, 115, 137));
        theme.getThemeFontScheme().getMajor().get(FontLanguageIndex.Latin).setName("Marion");
        theme.getThemeFontScheme().getMinor().get(FontLanguageIndex.Latin).setName("Marion");

        //Apply the above custom theme.
        workbook.setTheme(theme);

        //Set active cell.
        worksheet.getRange("B7").activate();

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
