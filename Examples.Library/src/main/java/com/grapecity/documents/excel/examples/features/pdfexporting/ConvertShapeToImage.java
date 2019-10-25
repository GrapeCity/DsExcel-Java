package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.rendering.PDFRenderer;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Point;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Size;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConvertShapeToImage extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
		
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Add a rectangle
        IShape rectangle = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 0, 0, 191, 194);

        // Set soild fill for rectangle
        rectangle.getFill().getColor().setRGB(Color.GetOrangeRed());

        // Create a PrintManager.
        PrintManager printManager = new PrintManager();

        // Get the size of the shape.
        IRange topLeftCell = rectangle.getTopLeftCell();
        IRange bottomRightCell = rectangle.getBottomRightCell();
        IRange shapeRange = worksheet.getRange(topLeftCell.getRow(), topLeftCell.getColumn(),
                bottomRightCell.getRow() - topLeftCell.getRow() + 1,
                bottomRightCell.getColumn() - topLeftCell.getColumn() + 1);
        Size size = printManager.getSize(shapeRange);
        PDRectangle pageSize = new PDRectangle((float) size.getWidth(), (float) size.getHeight());

        // Create a pdf document.
        PDDocument doc = new PDDocument();
        PDPage page = new PDPage(pageSize);
        doc.addPage(page);

        // Draw the printArea to the specified location on the page.
        printManager.draw(doc, page, new Point(0, 0), shapeRange);

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
