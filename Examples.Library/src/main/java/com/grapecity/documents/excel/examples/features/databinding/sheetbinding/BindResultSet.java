package com.grapecity.documents.excel.examples.features.databinding.sheetbinding;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.GcMockResultSet;

public class BindResultSet extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        //Here in the demo, we use a mock class to generate instance of java.sql.ResultSet.
        //User who use data binding in product, must get instance of java.sql.ResultSet from the
        //related database connection.
        java.sql.ResultSet datasource = new GcMockResultSet(this.getResourceStream("score.csv"));

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Set AutoGenerateColumns as false
        worksheet.setAutoGenerateColumns(false);

        // Bind columns manually.
        worksheet.getRange("A:A").getEntireColumn().setBindingPath("ID");
        worksheet.getRange("B:B").getEntireColumn().setBindingPath("Name");
        worksheet.getRange("C:C").getEntireColumn().setBindingPath("Score");
        worksheet.getRange("D:D").getEntireColumn().setBindingPath("Team");

        // Set data source
        worksheet.setDataSource(datasource);
    }

    @Override
    public String[] getResources() {
        return new String[] { "score.csv" };
    }

    @Override
    public String[] getRefs() {
        return new String[] { "GcMockResultSet", "GcResultSetMetaData" };
    }
}
