package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class ProtectWorkbook extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/Medical office start-up expenses 1.xlsx");
        workbook.open(fileStream);

        //Protects the workbook with a password so that other users cannot view hidden worksheets, add, move, delete, hide, or rename worksheets.
        //The protection only happens when you open it with an Excel application.
        workbook.protect("Y6dh!et5");
    }
    
    @Override
    public boolean getCanDownload()
    {
    	 return true;
    }
    
    @Override
    public boolean getShowViewer()
    {
    	return false;
    }
    
    @Override
    public String[] getResources() {
        return new String[] { "xlsx/Medical office start-up expenses 1.xlsx" };
    }
}
