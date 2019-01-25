package com.grapecity.documents.excel.examples.exceltemplates;

import java.net.URL;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SimpleInvoice extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Load template file Simple invoice.xlsx from resource
        workbook.open(this.getResourceStream("xlsx/Simple invoice.xlsx"));

        IWorksheet worksheet = workbook.getActiveSheet();

        // fill some new items
        worksheet.getRange("E09:H09").setValue(new Object[]{"DD1-001", "Item 3", 5.60, 12});
        worksheet.getRange("E10:H10").setValue(new Object[]{"DD2-001", "Item 3", 8.5, 14});
        worksheet.getRange("E11:H11").setValue(new Object[]{"DD3-001", "Item 3", 9.6, 16});

    }


    @Override
    public String getTemplateName() {

        return "Simple invoice.xlsx";
    }

    @Override
    public boolean getShowViewer() {

        return false;
    }

    @Override
    public String[] getResources() {
        return new String[] {"xlsx/Simple invoice.xlsx"};
    }
}
