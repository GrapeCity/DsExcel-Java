package com.grapecity.documents.excel.examples.features.charts.series;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ExtendSeries extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:D4").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 50},
                {"Item2", 15, -36, 40},
                {"Item3", 52, 40, -30},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D4"), RowCol.Columns, true, true);

        worksheet.getRange("A12:D13").setValue(new Object[][]{
                {"Item5", 10, 20, -30},
                {"Item6", 20, 40, 80},
        });

        //add new data point to existing series.
        shape.getChart().getSeriesCollection().extend(worksheet.getRange("A12:D13"), RowCol.Columns, true);

    }

}
