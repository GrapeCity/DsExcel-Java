package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ColumnClusteredChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:D2").setValue(new Object[][]{
                {100, 200, 300, 400},
                {100, 200, 300, 400}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D2"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Column Clustered Chart");

    }

}
