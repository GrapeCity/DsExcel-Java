package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ColumnStacked100Chart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnStacked100, 250, 20, 360, 230);
        worksheet.getRange("A1:B6").setValue(new Object[][]{
                {1, 5},
                {2, 4},
                {3, 3},
                {4, 2},
                {5, 1},
                {5, 3}
        });

        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B6"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Column Stacked 100 Chart");
    }

}
