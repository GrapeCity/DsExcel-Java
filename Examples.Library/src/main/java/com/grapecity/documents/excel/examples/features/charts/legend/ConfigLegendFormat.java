package com.grapecity.documents.excel.examples.features.charts.legend;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ILegend;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigLegendFormat extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, 36, 27},
                {"Item3", 52, 70, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);

        shape.getChart().setHasLegend(true);
        //config legend font style
        ILegend legend = shape.getChart().getLegend();
        legend.getFont().setSize(12);
        legend.getFont().setName("Cooper Black");
        //config legend format
        legend.getFormat().getFill().getColor().setRGB(Color.GetLightGray());
        legend.getFormat().getLine().getColor().setRGB(Color.GetGray());

    }

    @Override
    public boolean getShowViewer() {

        return true;

    }
}
