package com.grapecity.documents.excel.examples.features.sparklines;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.SparkType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class UseDateAxis extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        Object data = new Object[][]
                {
                        {"Number", "Date", "Customer", "Description", "Trend", "0-30 Days", "30-60 Days", "60-90 Days", ">90 Days", "Amount"},
                        {"1001", new GregorianCalendar(2017, 4, 21), "Customer A", "Invoice 1001", null, 1200.15, 1916.18, 1105.23, 1806.53, null},
                        {"1002", new GregorianCalendar(2017, 2, 18), "Customer B", "Invoice 1002", null, 896.23, 1005.53, 1800.56, 1150.49, null},
                        {"1003", new GregorianCalendar(2017, 5, 15), "Customer C", "Invoice 1003", null, 827.63, 1009.23, 1869.23, 1002.56, null}
                };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("B2:K5").setValue(data);
        worksheet.getRange("B:K").setColumnWidth(15);

        worksheet.getTables().add(worksheet.getRange("B2:K5"), true);
        worksheet.getTables().get(0).getColumns().get(9).getDataBodyRange().setFormula("=SUM(Table1[@[0-30 Days]:[>90 Days]])");

        //create a new group of sparklines.
        worksheet.getRange("F3:F5").getSparklineGroups().add(SparkType.Line, "G3:J5");

        worksheet.getRange("G7:J7").setValue(new Object[]{new GregorianCalendar(2011, 11, 16), new GregorianCalendar(2011, 11, 17), new GregorianCalendar(2011, 11, 18), new GregorianCalendar(2011, 11, 19)});
        worksheet.getRange("F3").getSparklineGroups().get(0).setDateRange("G7:J7");
        worksheet.getRange("F3").getSparklineGroups().get(0).getAxes().getHorizontal().getAxis().setVisible(true);
        worksheet.getRange("F3").getSparklineGroups().get(0).getAxes().getHorizontal().getAxis().getColor().setColor(Color.GetGreen());

    }

}
