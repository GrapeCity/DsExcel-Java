package com.grapecity.documents.excel.examples.features.spreadjsfeatures;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class BackgroundColor extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

    	//Set background color
        workbook.getBookView().setBackColor(Color.GetLightSkyBlue());
        workbook.getBookView().setGrayAreaBackColor(Color.GetGray());
        
      //Set value to a cell
        IWorksheet worksheet = workbook.getActiveSheet();
        worksheet.getRange("H20").setValue("The text");

        //Set page options
        worksheet.getPageSetup().setPrintGridlines(true);
        worksheet.getPageSetup().setPrintHeadings(true);
    }

    @Override
	public boolean getSavePdf() {
		return true;
	}

	@Override
    public boolean getIsNew() {
        return true;
    }
}
