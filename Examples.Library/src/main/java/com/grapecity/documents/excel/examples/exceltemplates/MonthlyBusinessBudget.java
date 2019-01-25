package com.grapecity.documents.excel.examples.exceltemplates;

import java.net.URL;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisGroup;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.DisplayUnit;
import com.grapecity.documents.excel.drawing.GradientStyle;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.IChart;
import com.grapecity.documents.excel.drawing.IPoint;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.LineDashStyle;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MonthlyBusinessBudget extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Load template file Monthly business budget.xlsx from resource
        workbook.open(this.getResourceStream("xlsx/Monthly business budget.xlsx"));

        IWorksheet worksheet = workbook.getActiveSheet();

        // change table style to TableStyleMedium14
        ITable totalsTable = worksheet.getTables().get("TotalsTable");
        totalsTable.setTableStyle(workbook.getTableStyles().get("TableStyleMedium14"));

        // change chart type to column stacked
        IChart chart = worksheet.getShapes().get(0).getChart();
        chart.setChartType(ChartType.ColumnStacked);
        chart.getColumnGroups().get(0).setOverlap(100);

        // set a bigger font size for chart title
        chart.getChartTitle().getFont().setSize(24);
        chart.getChartTitle().getFont().setBold(true);

        // give a one color gradient to chart area
        chart.getChartArea().getFormat().getFill().oneColorGradient(GradientStyle.Horizontal, 1, 0);
        chart.getChartArea().getFormat().getFill().getGradientStops().get(0).getColor().setObjectThemeColor(ThemeColor.Accent6);
        chart.getChartArea().getFormat().getFill().getGradientStops().get(0).getColor().setBrightness(0.8);

        // make fill of plot area transparant
        chart.getPlotArea().getFormat().getFill().setTransparency(1);

        // config series1 of chart
        ISeries series1 = chart.getSeriesCollection().get(0);
        series1.setHasDataLabels(true);
        series1.getFormat().getFill().getColor().setObjectThemeColor(ThemeColor.Accent1);

        // give some formatting for the first point of series1
        IPoint point1 = series1.getPoints().get(0);
        point1.getFormat().getLine().getColor().setRGB(Color.GetBlack());
        point1.getFormat().getLine().setWeight(2);
        point1.getFormat().getLine().setDashStyle(LineDashStyle.Dash);

        // config series2 of chart
        ISeries series2 = chart.getSeriesCollection().get(1);
        series2.setHasDataLabels(true);
        series2.getFormat().getFill().getColor().setObjectThemeColor(ThemeColor.Accent6);
        series2.getDataLabels().getFont().getColor().setRGB(Color.GetRed());

        // get the value axis
        IAxis value_axis = chart.getAxes().item(AxisType.Value, AxisGroup.Primary);

        // show the display unit as thousands for value axis
        value_axis.setHasDisplayUnitLabel(true);
        value_axis.setDisplayUnit(DisplayUnit.Thousands);

        // give a color for the major grid line of value axis
        value_axis.getMajorGridlines().getFormat().getLine().getColor().setObjectThemeColor(ThemeColor.Accent6);

    }

    @Override
    public String getTemplateName() {

        return "Monthly business budget.xlsx";
    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public String[] getResources() {
        return new String[] {"xlsx/Monthly business budget.xlsx"};
    }
}
