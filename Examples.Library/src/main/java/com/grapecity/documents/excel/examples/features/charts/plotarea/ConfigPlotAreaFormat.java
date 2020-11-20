package com.grapecity.documents.excel.examples.features.charts.plotarea;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IPlotArea;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigPlotAreaFormat extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, 36, 27},
                {"Item3", 52, 50, -30},
                {"Item4", 22, 65, 30},
                {"Item5", 23, 40, 69}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);

        IPlotArea plotarea = shape.getChart().getPlotArea();
        plotarea.getFormat().getFill().getColor().setRGB(Color.GetLightGray());
        plotarea.getFormat().getLine().getColor().setRGB(Color.GetGray());
        plotarea.getFormat().getLine().setWeight(1);

    }

}
