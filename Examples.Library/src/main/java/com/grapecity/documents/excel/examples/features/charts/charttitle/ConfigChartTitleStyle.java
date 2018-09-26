package com.grapecity.documents.excel.examples.features.charts.charttitle;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigChartTitleStyle extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -20, 36, 27},
                {"Item3", 62, 70, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 50, 50}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);

        //config chart title style
        shape.getChart().setHasTitle(true);
        shape.getChart().getChartTitle().setText("MyChartTitle");
        shape.getChart().getChartTitle().getFont().setSize(15);
        shape.getChart().getChartTitle().getFont().setName("Bradley Hand ITC");
        shape.getChart().getChartTitle().getFormat().getFill().getColor().setRGB(Color.GetDarkOrange());
        shape.getChart().getChartTitle().getFormat().getLine().getColor().setRGB(Color.GetCornflowerBlue());
        shape.getChart().getChartTitle().getFormat().getLine().setWeight(1.5);

    }

    @Override
    public boolean getShowViewer() {

        return true;

    }

}
