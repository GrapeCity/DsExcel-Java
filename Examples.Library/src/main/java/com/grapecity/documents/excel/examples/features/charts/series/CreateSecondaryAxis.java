package com.grapecity.documents.excel.examples.features.charts.series;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisGroup;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateSecondaryAxis extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:C6").setValue(new Object[][]{
                {null, "S1", "S2"},
                {"Item1", 10, 25},
                {"Item2", -51, -36},
                {"Item3", 32, 64},
                {"Item4", 44, 80},
                {"Item5", 60, 100}
        });

        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C6"), RowCol.Columns, true, true);
        ISeries series2 = shape.getChart().getSeriesCollection().get(1);
        //add a secondary axis
        series2.setAxisGroup(AxisGroup.Secondary);
        series2.setChartType(ChartType.Line);

    }

}
