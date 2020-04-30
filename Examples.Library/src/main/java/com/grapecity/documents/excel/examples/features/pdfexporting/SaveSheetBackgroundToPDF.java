package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PdfSaveOptions;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SaveSheetBackgroundToPDF extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").setValue("GrapeCity Documents for Excel");
        worksheet.getRange("A1").getFont().setSize(25);
        
        //Set a background image for worksheet
        InputStream inputStream = this.getResourceStream("logo.png");
        try {
            byte[] bytes = new byte[inputStream.available()];
            inputStream.read(bytes, 0, bytes.length);
            worksheet.setBackgroundPicture(bytes);
        }catch (IOException ioe){
            ioe.printStackTrace();
        }

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        //Print the background image when saving pdf.
        //The background image will be centered on every page of the sheet.
        pdfSaveOptions.setPrintBackgroundPicture(true);

        //Save the workbook into pdf file.
        workbook.save(outputStream, pdfSaveOptions);
	}

    @Override
	public boolean getSavePageInfos() {
        return true;
	}
	
    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public String[] getResources() {
        return new String[]{"logo.png"};
    }
}
