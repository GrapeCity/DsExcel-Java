package com.grapecity.documents.excel.examples.features.charts.markers;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.LineStyle;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigMarkersFormat extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.LineMarkers, 250, 20, 360, 230);
        worksheet.getRange("A1:B6").setValue(new Object[][]{
                {null, "S1"},
                {"Item1", 10},
                {"Item2", -51},
                {"Item3", 52},
                {"Item4", 22},
                {"Item5", 40}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);
        ISeries series1 = shape.getChart().getSeriesCollection().get(0);

        //config line markers style
        series1.getMarkerFormat().getFill().getColor().setRGB(Color.GetCornflowerBlue());
        series1.getMarkerFormat().getLine().setStyle(LineStyle.ThickThin);
        series1.getMarkerFormat().getLine().getColor().setRGB(Color.GetLightGreen());
        series1.getMarkerFormat().getLine().setWeight(3);

    }

}
