package com.grapecity.documents.excel.examples.features.pdfexporting.exportshape;

import java.io.InputStream;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CustomShape extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
        //Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/CustomShapes.xlsx");
        workbook.open(fileStream);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
    
    @Override
    public boolean getShowScreenshot() {
    	return true;
    }
    
    @Override
    public String[] getResources() {
        return new String[]{"xlsx/CustomShapes.xlsx"};
    }
    
}
