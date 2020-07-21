package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.SheetType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddChartSheet extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        worksheet.getRange("A1:D6").setValue(new Object[][]{
            {null, "S1", "S2", "S3"},
            {"Item1", 10, 25, 25},
            {"Item2", 51, 36, 27},
            {"Item3", 52, 85, 30},
            {"Item4", 22, 65, 65},
            {"Item5", 23, 69, 69}
        });

        //Add a chart sheet.
        IWorksheet chartSheet = workbook.getWorksheets().add(SheetType.Chart);

        //Add the main chart for the chart sheet.
        IShape mainChart = chartSheet.getShapes().addChart(ChartType.ColumnClustered, 100, 100, 200, 200);
        mainChart.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"));

        //Make the chart sheet the active sheet.
        chartSheet.activate();
    }
    
    @Override
    public boolean getShowViewer()
    {
    	return false;
    }
}
