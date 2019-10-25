package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.rendering.PDFRenderer;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Point;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Size;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConvertSheetToImage extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
		
        workbook.open(this.getResourceStream("xlsx/Employee absence schedule.xlsx"));
        
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Create a PrintManager.
        PrintManager printManager = new PrintManager();

        //Get the first print area of the worksheet.
        IRange printArea = printManager.getPrintAreas(worksheet).get(0);
        //Get the size of the printArea.
        Size size = printManager.getSize(printArea);
        PDRectangle pageSize = new PDRectangle((float)size.getWidth(), (float)size.getHeight());

        //Create a pdf document.
        PDDocument doc = new PDDocument();
        PDPage page = new PDPage(pageSize);
        doc.addPage(page);	

        //Draw the print to the specified location on the page. 
        printManager.draw(doc, page, new Point(0, 0), printArea);

        //Saves the page as an image to a stream.
        PDFRenderer pdfRenderer = new PDFRenderer(doc);
        BufferedImage bim;
		try {
        	bim = pdfRenderer.renderImage(0);
			ImageIO.write(bim, "PNG", outputStream); 
			doc.close();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
    }

    @Override
    public String getTemplateName() {
        return "Employee absence schedule.xlsx";
    }

    @Override
    public boolean getSaveAsImage() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public String[] getResources() {
        return new String[] {"xlsx/Employee absence schedule.xlsx"};
    }
    
}