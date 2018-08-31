package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class SaveSparklinesToPDF extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        Object[][] data = new Object[][]
            {
                {"Customer", "0-30 Days", "30-60 Days", "60-90 Days", ">90 Days"},
                {"Customer A", 1200.15, 1916.18, 1105.23, 1806.53},
                {"Customer B", 896.23, 1005.53, 1800.56, 1150.49},
                {"Customer C", 827.63, 1009.23, 1869.23, 1002.56}
            };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("B2:E5").setValue(data);
        worksheet.getRange("B:F").setColumnWidth(15);
        worksheet.getRange("B:E").setHorizontalAlignment(HorizontalAlignment.Center);
        ITable table = worksheet.getTables().add(worksheet.getRange("B2:F5"), true);
        table.setTableStyle(workbook.getTableStyles().get("TableStyleMedium3"));
        table.getColumns().get(4).setName("Sparklines");

        //create a new group of sparklines.
        worksheet.getRange("F3").getSparklineGroups().add(SparkType.Line, "C3:E3");
        worksheet.getRange("F4").getSparklineGroups().add(SparkType.Column, "C4:E4");
        worksheet.getRange("F5").getSparklineGroups().add(SparkType.ColumnStacked100, "C5:E5");
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

}