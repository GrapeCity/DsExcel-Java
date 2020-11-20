package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisGroup;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CombinationChart2 extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:C17").setValue(new Object[][]{
                {"Area 1", "Column 1", "Column 2"},
                {1350, 120, 75},
                {1500, 90, 35},
                {1200, 80, 50},
                {1300, 80, 80},
                {1750, 90, 100},
                {1640, 120, 130},
                {1700, 120, 95},
                {1100, 90, 80},
                {1350, 120, 75},
                {1500, 90, 35},
                {1200, 80, 50},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C17"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Combination Chart");
        ISeries series1 = shape.getChart().getSeriesCollection().get(0);
        ISeries series2 = shape.getChart().getSeriesCollection().get(1);
        ISeries series3 = shape.getChart().getSeriesCollection().get(2);
        //change series type
        series1.setChartType(ChartType.Area);
        series2.setChartType(ChartType.ColumnStacked);
        series3.setChartType(ChartType.ColumnStacked);
        //set axis group
        series2.setAxisGroup(AxisGroup.Secondary);
        series3.setAxisGroup(AxisGroup.Secondary);
        //config axis scale and unit
        IAxis value_axis = shape.getChart().getAxes().item(AxisType.Value);
        IAxis value_second_axis = shape.getChart().getAxes().item(AxisType.Value, AxisGroup.Secondary);
        value_axis.setMaximumScale(1800);
        value_axis.setMajorUnit(450);
        value_second_axis.setMaximumScale(300);
        value_second_axis.setMajorUnit(75);

    }

}
