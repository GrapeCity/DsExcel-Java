package com.grapecity.documents.excel.examples.excelreporting;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITheme;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Theme;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.TotalsCalculation;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisGroup;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.GradientStyle;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.PresetGradientType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CostAnalysisWithParetoChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        Object data = new Object[][]{
                {"Cost Center", "Annual Cost", "Percent of Total", "Cumulative Percent"},
                {"Parts and materials", 1325000, null, null},
                {"Manufacturing equipment", 900500, null, null},
                {"Salaries", 575000, null, null},
                {"Maintenance", 395000, null, null},
                {"Office lease", 295000, null, null},
                {"Warehouse lease", 250000, null, null},
                {"Insurance", 180000, null, null},
                {"Benefits and pensions", 130000, null, null},
                {"Vehicles", 125000, null, null},
                {"Research", 75000, null, null},
        };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.setName("COST DATA and CHART");
        worksheet.setTabColor(Color.FromArgb(63, 94, 101));
        worksheet.getSheetView().setDisplayGridlines(false);

        //Set Value.
        worksheet.getRange("B2").setValue("COST ANALYSIS - PARETO");
        worksheet.getRange("B39").setValue("COST ANALYSIS");
        worksheet.getRange("B41:E51").setValue(data);

        //Set NumberFormat.
        worksheet.getRange("C42:C52").setNumberFormat("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)");
        worksheet.getRange("D42:E52").setNumberFormat("0.00%");

        //Change the range's RowHeight and ColumnWidth.
        worksheet.setStandardHeight(15.75);
        worksheet.setStandardWidth(8.43);
        worksheet.getRange("1:1").setRowHeight(12.75);
        worksheet.getRange("2:2").setRowHeight(20.25);
        worksheet.getRange("3:3").setRowHeight(16.5);
        worksheet.getRange("39:39").setRowHeight(20.25);
        worksheet.getRange("40:40").setRowHeight(16.5);
        worksheet.getRange("41:41").setRowHeight(29.25);
        worksheet.getRange("42:51").setRowHeight(20.1);
        worksheet.getRange("A:A").setColumnWidth(1.44140625);
        worksheet.getRange("B:B").setColumnWidth(25.88671875);
        worksheet.getRange("C:C").setColumnWidth(18.5546875);
        worksheet.getRange("D:D").setColumnWidth(17.77734375);
        worksheet.getRange("E:E").setColumnWidth(20.77734375);

        //Apply one build in name style on the ranges.
        worksheet.getRange("B2:J2").setStyle(workbook.getStyles().get("Heading 1"));
        worksheet.getRange("B39:E39").setStyle(workbook.getStyles().get("Heading 1"));

        //Create a table and apply one build in table style.
        ITable table = worksheet.getTables().add(worksheet.getRange("B41:E51"), true);
        table.setName("tblData");
        table.setTableStyle(workbook.getTableStyles().get("TableStyleLight9"));
        table.setShowTotals(true);
        table.setShowTableStyleRowStripes(true);

        //Use table formula in table range.
        worksheet.getRange("D42:D51").setFormula("=[@[Annual Cost]]/SUM([[Annual Cost]])");
        worksheet.getRange("E42:E51").setFormula("=SUM(INDEX([Percent of Total],1):[@[Percent of Total]])");
        table.getColumns().get(1).setTotalsCalculation(TotalsCalculation.Sum);
        table.getColumns().get(2).setTotalsCalculation(TotalsCalculation.Sum);
        table.getColumns().get(3).setTotalsCalculation(TotalsCalculation.None);

        //Add chart.
        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 9.75, 48, 597, 472.5);
        shape.setName("Pareto Chart");

        //Add Series.
        ISeries series_ColumnClustered = shape.getChart().getSeriesCollection().newSeries();
        series_ColumnClustered.setFormula("=SERIES('COST DATA and CHART'!$C$41,'COST DATA and CHART'!$B$42:$B$51,'COST DATA and CHART'!$C$42:$C$51,1)");
        series_ColumnClustered.getFormat().getFill().twoColorGradient(GradientStyle.Horizontal, 1);
        series_ColumnClustered.getFormat().getFill().setGradientAngle(90);
        series_ColumnClustered.getFormat().getFill().getGradientStops().get(0).getColor().setObjectThemeColor(ThemeColor.Accent1);
        series_ColumnClustered.getFormat().getFill().getGradientStops().get(0).getColor().setBrightness(0);
        series_ColumnClustered.getFormat().getFill().getGradientStops().get(0).setPosition(0);
        series_ColumnClustered.getFormat().getFill().getGradientStops().get(1).getColor().setObjectThemeColor(ThemeColor.Accent1);
        series_ColumnClustered.getFormat().getFill().getGradientStops().get(1).getColor().setBrightness(-0.16);
        series_ColumnClustered.getFormat().getFill().getGradientStops().get(1).setPosition(1);
        series_ColumnClustered.getFormat().getLine().getColor().setObjectThemeColor(ThemeColor.Light1);

        ISeries series_Line = shape.getChart().getSeriesCollection().newSeries();
        series_Line.setChartType(ChartType.Line);
        series_Line.setFormula("=SERIES('COST DATA and CHART'!$E$41,,'COST DATA and CHART'!$E$42:$E$51,2)");
        series_Line.getFormat().getLine().setWeight(2.25);
        series_Line.setAxisGroup(AxisGroup.Secondary);

        //Change the secondary's maxinumscale.
        IAxis secondary_axis = shape.getChart().getAxes().item(AxisType.Value, AxisGroup.Secondary);
        secondary_axis.setMaximumScale(1.0);

        //Set the chart's title format.
        shape.getChart().getChartTitle().setText("Cost Center");
        shape.getChart().getChartTitle().getFont().setThemeFont(ThemeFont.Minor);
        shape.getChart().getChartTitle().getFont().getColor().setRGB(Color.FromArgb(89, 89, 89));
        shape.getChart().getChartTitle().getFont().setSize(18);

        //Set the chart has no legend.
        shape.getChart().setHasLegend(false);

        //Set the char group's Overlap and GapWidth.
        shape.getChart().getColumnGroups().get(0).setOverlap(0);
        shape.getChart().getColumnGroups().get(0).setGapWidth(0);

        //Set chart area's format.
        shape.getChart().getChartArea().getFormat().getFill().presetGradient(GradientStyle.Horizontal, 1, PresetGradientType.EarlySunset);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().delete(3);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().delete(3);
        shape.getChart().getChartArea().getFormat().getFill().setGradientAngle(90);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().get(0).getColor().setObjectThemeColor(ThemeColor.Light1);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().get(0).getColor().setBrightness(0);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().get(0).setPosition(0);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().get(1).getColor().setObjectThemeColor(ThemeColor.Light1);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().get(1).getColor().setBrightness(-0.15);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().get(1).setPosition(0.68);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().get(2).getColor().setObjectThemeColor(ThemeColor.Light1);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().get(2).getColor().setBrightness(0);
        shape.getChart().getChartArea().getFormat().getFill().getGradientStops().get(2).setPosition(1);

        //Create customize theme.
        ITheme theme = new Theme("test");
        theme.getThemeColorScheme().get(ThemeColor.Dark1).setRGB(Color.FromArgb(0, 0, 0));
        theme.getThemeColorScheme().get(ThemeColor.Light1).setRGB(Color.FromArgb(255, 255, 255));
        theme.getThemeColorScheme().get(ThemeColor.Dark2).setRGB(Color.FromArgb(96, 89, 88));
        theme.getThemeColorScheme().get(ThemeColor.Light2).setRGB(Color.FromArgb(241, 246, 246));
        theme.getThemeColorScheme().get(ThemeColor.Accent1).setRGB(Color.FromArgb(63, 94, 101));
        theme.getThemeColorScheme().get(ThemeColor.Accent2).setRGB(Color.FromArgb(224, 170, 83));
        theme.getThemeColorScheme().get(ThemeColor.Accent3).setRGB(Color.FromArgb(179, 29, 66));
        theme.getThemeColorScheme().get(ThemeColor.Accent4).setRGB(Color.FromArgb(162, 67, 162));
        theme.getThemeColorScheme().get(ThemeColor.Accent5).setRGB(Color.FromArgb(120, 59, 101));
        theme.getThemeColorScheme().get(ThemeColor.Accent6).setRGB(Color.FromArgb(55, 120, 169));
        theme.getThemeColorScheme().get(ThemeColor.Hyperlink).setRGB(Color.FromArgb(71, 166, 181));
        theme.getThemeColorScheme().get(ThemeColor.FollowedHyperlink).setRGB(Color.FromArgb(120, 59, 101));
        theme.getThemeFontScheme().getMajor().get(FontLanguageIndex.Latin).setName("Constantia");
        theme.getThemeFontScheme().getMinor().get(FontLanguageIndex.Latin).setName("Helvetica");

        //Apply the above custom theme.
        workbook.setTheme(theme);

        //Set active cell.
        worksheet.getRange("B43").activate();

    }


    @Override
    public boolean getShowViewer() {

        return false;

    }

}
