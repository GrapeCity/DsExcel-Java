package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.rendering.PDFRenderer;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Point;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Size;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConvertRangeToImage extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A1:C1").setValue(new String[] { "Device", "Quantity", "Unit Price" });
        worksheet.getRange("A2:C5").setValue(new Object[][] 
        { 
            { "T540p", 12, 9850 }, 
            { "T570", 5, 7460 },
            { "Y460", 6, 5400 },
            { "Y460F", 8, 6240 } 
         });

        // Set style
        worksheet.getRange("A1:C1").getFont().setBold(true);
        worksheet.getRange("A1:C1").getFont().setColor(Color.GetWhite());
        worksheet.getRange("A1:C1").getInterior().setColor(Color.GetLightBlue());
        worksheet.getRange("A2:C5").getBorders().get(BordersIndex.InsideHorizontal).setColor(Color.GetOrange());
        worksheet.getRange("A2:C5").getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.DashDot);

        // Create a PrintManager.
        PrintManager printManager = new PrintManager();

        // Get the size of the printArea.
        Size size = printManager.getSize(worksheet.getRange("A1:C5"));
        PDRectangle pageSize = new PDRectangle((float) size.getWidth(), (float) size.getHeight());

        // Create a pdf document.
        PDDocument doc = new PDDocument();
        PDPage page = new PDPage(pageSize);
        doc.addPage(page);

        // Draw the Range"A1:C5" to the specified location on the page.
        printManager.draw(doc, page, new Point(0, 0), worksheet.getRange("A1:C5"));

        // Saves the page as an image to a stream.
        PDFRenderer pdfRenderer = new PDFRenderer(doc);
        BufferedImage bim;
        try {
        	bim = pdfRenderer.renderImage(0);
        	ImageIO.write(bim, "PNG", outputStream);
        	doc.close();
        } catch (IOException e) {
        	e.printStackTrace();
        }
	}

	@Override
	public boolean getSaveAsImage() {
        return true;
	}

	@Override
	public boolean getShowViewer() {
        return false;
	}

}
