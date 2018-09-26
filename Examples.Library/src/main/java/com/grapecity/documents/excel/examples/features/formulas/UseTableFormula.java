package com.grapecity.documents.excel.examples.features.formulas;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class UseTableFormula extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A:E").setColumnWidth(15);
        worksheet.getRange("A1:E3").setValue(new Object[][]{
                {"SalesPerson", "Region", "SalesAmount", "ComPct", "ComAmt"},
                {"Joe", "North", 260, 0.10, null},
                {"Nia", "South", 660, 0.15, null},
        });

        worksheet.getTables().add(worksheet.getRange("A1:E3"), true);
        worksheet.getTables().get(0).setName("DeptSales");
        worksheet.getTables().get(0).getColumns().get("ComPct").getDataBodyRange().setNumberFormat("0%");

        //Use table formula in table range.
        worksheet.getTables().get(0).getColumns().get("ComAmt").getDataBodyRange().setFormula("=[@ComPct]*[@SalesAmount]");

        //Use table formula out of table range.
        worksheet.getRange("F2").setFormula("=SUM(DeptSales[@SalesAmount])");
        worksheet.getRange("G2").setFormula("=SUM(DeptSales[[#Data],[SalesAmount]])");
        worksheet.getRange("H2").setFormula("=SUM(DeptSales[SalesAmount])");
        worksheet.getRange("I2").setFormula("=SUM(DeptSales[@ComPct], DeptSales[@ComAmt])");

        //judge if Range F2:I2 have formula.
        for (int i = 5; i <= 8; i++) {
            if (worksheet.getRange(1, i).getHasFormula()) {
                worksheet.getRange(1, i).getInterior().setColor(Color.GetLightBlue());
            }
        }
    }

}
