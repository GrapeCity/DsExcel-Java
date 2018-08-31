package com.grapecity.documents.excel.examples.features.charts.chartlines;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.drawing.TrendlineType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateTrendline extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);

        ISeries series1 = shape.getChart().getSeriesCollection().get(0);

        series1.getTrendlines().add();
        series1.getTrendlines().get(0).setType(TrendlineType.Linear);
        series1.getTrendlines().get(0).setForward(1);
        series1.getTrendlines().get(0).setBackward(0.5);
        series1.getTrendlines().get(0).setIntercept(2.5);
        series1.getTrendlines().get(0).setDisplayEquation(true);
        series1.getTrendlines().get(0).setDisplayRSquared(true);

    }


    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public boolean getShowScreenshot() {

        return true;

    }

}
