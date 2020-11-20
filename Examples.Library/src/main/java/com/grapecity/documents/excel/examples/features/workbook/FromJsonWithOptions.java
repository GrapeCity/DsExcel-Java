package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.DeserializationOptions;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FromJsonWithOptions extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("B3:C16").setValue(new Object[][] {
            { "ITEM", "AMOUNT" },
            { "Rent/mortgage", 800 },
            { "Electric", 120 },
            { "Gas", 50 },
            { "Cell phone", 45 },
            { "Groceries", 500 },
            { "Car payment", 273 },
            { "Auto expenses", 120 },
            { "Student loans", 50 },
            { "Credit cards", 100 },
            { "Auto Insurance", 78 },
            { "Personal care", 50 },
            { "Entertainment", 100 },
            { "Miscellaneous", 50 },
        });

        // Create a table
        ITable expensesTable = worksheet.getTables().add(worksheet.getRange("B3:C16"), true);
        expensesTable.setName("tblExpenses");
        worksheet.getRange("C3:C16").setNumberFormat("$#,##0_);($#,##0)");

        worksheet.getRange("B2:C2").merge();
        worksheet.getRange("B2").setValue("MONTHLY EXPENSES");
        worksheet.getRange("B2").getInterior().setColor(Color.FromArgb(219, 219, 219));
        worksheet.getRange("E2").setValue("Total Monthly Expenses");
        worksheet.getRange("E3").setFormula("SUM(tblExpenses[AMOUNT])");
        worksheet.getRange("E3").setNumberFormat("$#,##0_);($#,##0)");

        worksheet.getRange("B:B").setColumnWidth(15);
        worksheet.getRange("C:C").setColumnWidth(15);
        worksheet.getRange("E:F").setColumnWidth(15);

        String json = workbook.toJson();

        // ignore style and formula when deserialize workbook from json.
        DeserializationOptions deserializationOptions = new DeserializationOptions();
        deserializationOptions.setIgnoreStyle(true);
        deserializationOptions.setIgnoreFormula(true);
        workbook.fromJson(json, deserializationOptions);
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
