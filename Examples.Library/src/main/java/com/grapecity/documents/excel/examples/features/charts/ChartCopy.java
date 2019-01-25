package com.grapecity.documents.excel.examples.features.charts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ChartCopy extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        //Create chart, chart's range is Range["G1:M21"]
        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 300, 10, 300, 300);
        worksheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
        });

        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);

        //Range["G1:M21"] must contain chart's range, copy a new shape to Range["N1:T21"]
        worksheet.getRange("G1:M21").copy(worksheet.getRange("N1"));
        //worksheet.Range["G1:M21"].Copy(worksheet.Range["N1:T21"]);

        //Cross sheet copy, copy a new chart to worksheet2's Range["N1:T21"]
        //IWorksheet worksheet2 = workbook.Worksheets.Add()
        //worksheet.Range["G1:M21"].Copy(worksheet2.Range["E1"]);
        //worksheet.Range["G1:M21"].Copy(worksheet2.Range["N1:T21"]);

    }
}
