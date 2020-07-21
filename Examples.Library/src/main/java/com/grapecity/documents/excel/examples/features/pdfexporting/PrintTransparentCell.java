package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PdfSaveOptions;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IBackgroundPicture;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PrintTransparentCell extends ExampleBase {
	@Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream){
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Initialize worksheet's values.
        worksheet.getRange("A1").setValue("Info from Acme Institute of Health:");
        worksheet.getRange("B2").setValue("BLOOD PRESSURE TRACKER");
        worksheet.getRange("B4:F13").setValue(new Object[][] {
            { "NAME", null, null, null, "JAMES HILL" },
            { null, null, null, null, null },
            { null, null, null, "Systolic", "Diastolic" },
            { "TARGET BLOOD PRESSURE", null, null, 120, 80 },
            { null, null, null, null, null },
            { null, null, null, "Systolic", "Diastolic" },
            { "CALL PHYSICIAN IF ABOVE", null, null, 140, 90 },
            { null, null, null, null, null },
            { null, null, null, null, null },
            { "PHYSICIAN PHONE NUMBER", null, null, "(001))5104234242", null}
        });
        worksheet.getRange("A1").getFont().setSize(25);
        
      //Set row height.
        worksheet.setStandardHeight(12.75);
        worksheet.setStandardWidth(8.43);
        worksheet.getRows().get(1).setRowHeight(32.25);
        worksheet.getRows().get(2).setRowHeight(13.5);
        worksheet.getRows().get(3).setRowHeight(18.75);
        worksheet.getRows().get(6).setRowHeight(18.75);
        worksheet.getRows().get(9).setRowHeight(18.75);
        worksheet.getRows().get(12).setRowHeight(18.75);
        worksheet.getRows().get(15).setRowHeight(19.5);
        worksheet.getRows().get(16).setRowHeight(13.5);
        worksheet.getRows().get(33).setRowHeight(19.5);
        worksheet.getRows().get(34).setRowHeight(13.5);

        //Set column width.
        worksheet.getColumns().get(0).setColumnWidth(1.7109375);
        worksheet.getColumns().get(1).setColumnWidth(12.140625);
        worksheet.getColumns().get(2).setColumnWidth(12.140625);
        worksheet.getColumns().get(3).setColumnWidth(12.140625);
        worksheet.getColumns().get(4).setColumnWidth(11.85546875);
        worksheet.getColumns().get(5).setColumnWidth(12.7109375);

        // Set the transparency value of the background color of range["A1:G20"] to 50.
        worksheet.getRange("A1:G20").getInterior().setColor(Color.FromArgb(50, 255, 0, 0));

        // Add a background picture.
        InputStream stream = this.getResourceStream("AcmeLogo-vertical-250px.png");
        IBackgroundPicture picture = null;
        try {
        	picture = worksheet.getBackgroundPictures().addPictureInPixel(stream, ImageType.PNG, 10, 10, 300, 150);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        // You must create a pdfSaveOptions object before using PrintTransparentCell.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set print the transparency of cell's background color, so the background picture will come out in the back.
        pdfSaveOptions.setPrintTransparentCell(true);
        
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
        return new String[] { "AcmeLogo-vertical-250px.png" };
	}
}
