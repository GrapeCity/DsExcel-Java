package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ProtectWorkbook extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").setValue("GrapeCity Documents for Excel");

        //Protects the workbook so that other users cannot view hidden worksheets, add, move, delete, hidie, or rename worksheets.
        //The protection only happens when you open it with an Excel application.
        workbook.protect();
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
    public boolean getIsNew() {

        return true;

    }
}
