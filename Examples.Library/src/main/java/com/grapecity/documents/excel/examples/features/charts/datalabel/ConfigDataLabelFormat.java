package com.grapecity.documents.excel.examples.features.charts.datalabel;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.chart.ChartType;
import com.grapecity.documents.excel.drawing.chart.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.style.color.Color;

public class ConfigDataLabelFormat extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:B5").setValue(new Object[][]{
                {null, "S1"},
                {"Item1", -20},
                {"Item2", 30},
                {"Item3", 50},
                {"Item3", 40}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B5"), RowCol.Columns, true, true);
        ISeries series1 = shape.getChart().getSeriesCollection().get(0);
        series1.setHasDataLabels(true);
        series1.getDataLabels().setShowSeriesName(true);

        //set series1's all data label's format.
        series1.getDataLabels().getFormat().getFill().getColor().setRGB(Color.getPink());
        series1.getDataLabels().getFormat().getLine().getColor().setRGB(Color.getGreen());
        series1.getDataLabels().getFormat().getLine().setWeight(1);

        //set series1's specific data label's format.
        series1.getDataLabels().get(2).getFormat().getFill().getColor().setRGB(Color.getLightGreen());
        series1.getPoints().get(2).getDataLabel().getFormat().getLine().getColor().setRGB(Color.getGray());
        series1.getPoints().get(2).getDataLabel().getFormat().getLine().setWeight(2);

    }

}
