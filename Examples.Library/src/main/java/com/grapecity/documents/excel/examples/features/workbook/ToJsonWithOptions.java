package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.SerializationOptions;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ToJsonWithOptions extends ExampleBase {
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

        // ignore style and formula when serialize workbook to json.
        SerializationOptions serializationOptions = new SerializationOptions();
        serializationOptions.setIgnoreStyle(true);
        serializationOptions.setIgnoreFormula(true);

        String jsonWithOption = workbook.toJson(serializationOptions);

        workbook.fromJson(jsonWithOption);
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
