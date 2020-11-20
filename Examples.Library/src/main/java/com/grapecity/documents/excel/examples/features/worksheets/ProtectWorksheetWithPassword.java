package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class ProtectWorksheetWithPassword extends ExampleBase {
	public void execute(Workbook workbook) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/Medical office start-up expenses 1.xlsx");
        workbook.open(fileStream);

        //Use a password to protect all the worksheet. If you forget the password, you cannot unprotect the worksheet.
        for (IWorksheet worksheet : workbook.getWorksheets())
        {
            worksheet.protect("Y6dh!et5");
        }
	}

	@Override
	public boolean getShowViewer() {

        return false;

	}

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/Medical office start-up expenses 1.xlsx" };
    }
}
