package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AreaChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.Area, 250, 20, 360, 230);
        worksheet.getRange("A1:C13").setValue(new Object[][]{
                {null, "Blue Series", "Orange Series"},
                {"Jan", 0, 59.1883603948205},
                {"Feb", 44.6420211591501, 52.2280901938606},
                {"Mar", 45.2174930051225, 49.8093056416248},
                {"Apr", 62, 37.3065749226828},
                {"May", 53, 34.4312192530766},
                {"Jun", 31.8933622049831, 69.7834561753736},
                {"Jul", 41.7930895085093, 63.9418103906982},
                {"Aug", 73, 57.4049534494926},
                {"Sep", 49.8773891668518, 33},
                {"Oct", 50, 74},
                {"Nov", 54.7658428630216, 22.9587876597096},
                {"Dec", 32, 54},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C13"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Area Chart");
    }

}
