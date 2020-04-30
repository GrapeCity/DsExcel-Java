package com.grapecity.documents.excel.examples.features.pdfexporting.printmanager;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Point;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class DrawRangeInsideAPDF extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A4:C4").setValue(new String[] {"Device", "Quantity", "Unit Price"});
        worksheet.getRange("A5:C8").setValue(new Object[][]
        {
        	{"T540p", 12, 9850},
        	{"T570", 5, 7460},
        	{"Y460", 6, 5400},
        	{"Y460F", 8, 6240}
        });

        //Set style
        worksheet.getRange("A4:C4").getFont().setBold(true);
        worksheet.getRange("A4:C4").getFont().setColor(Color.GetWhite());
        worksheet.getRange("A4:C4").getInterior().setColor(Color.GetLightBlue());
        worksheet.getRange("A5:C8").getBorders().get(BordersIndex.InsideHorizontal).setColor(Color.GetOrange());
        worksheet.getRange("A5:C8").getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.DashDot);
        
        //Create a pdf document.
        PDDocument doc = new PDDocument();
        PDPage page = new PDPage();
        doc.addPage(page);	

        //Create a PrintManager.
        PrintManager printManager = new PrintManager();

        // Draw the Range"A4:C8" to the specified location on the page. 
        printManager.draw(doc, page, new Point(30, 100), worksheet.getRange("A4:C8"));

        //Save the modified pages into pdf file.
        try {
            doc.save(outputStream);
            doc.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
	}

	@Override
	public boolean getSavePageInfos() {
        return true;
	}

	@Override
	public boolean getShowViewer() {
        return false;
	}
}
