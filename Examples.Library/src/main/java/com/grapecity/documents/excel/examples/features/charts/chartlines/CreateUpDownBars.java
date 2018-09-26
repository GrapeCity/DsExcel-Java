package com.grapecity.documents.excel.examples.features.charts.chartlines;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateUpDownBars extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.Line, 250, 20, 360, 230);
        worksheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);

        //config up down bars for line chart.
        shape.getChart().getLineGroups().get(0).setHasUpDownBars(true);
        shape.getChart().getLineGroups().get(0).getUpBars().getFormat().getFill().getColor().setRGB(Color.FromArgb(199, 235, 217));
        shape.getChart().getLineGroups().get(0).getDownBars().getFormat().getFill().getColor().setRGB(Color.GetLightPink());

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
