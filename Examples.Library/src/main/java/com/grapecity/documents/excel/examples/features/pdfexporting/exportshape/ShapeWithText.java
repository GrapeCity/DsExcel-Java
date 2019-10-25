package com.grapecity.documents.excel.examples.features.pdfexporting.exportshape;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ITextRange;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShapeWithText extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
		
        IWorksheet sheet = workbook.getWorksheets().get(0);
        
        // Add a rectangle
        IShape rectangle = sheet.getShapes().addShape(AutoShapeType.Rectangle, 50, 30, 300, 200);
        
        // Add rich text to rectangle
        rectangle.getFill().getColor().setRGB(Color.GetWhite());

        // Add first paragraph
        ITextRange run1 = rectangle.getTextFrame().getTextRange().getParagraphs().get(0).getRuns().add("         Doc");
        run1.getFont().getColor().setRGB(Color.GetTomato());
        ITextRange run2 = rectangle.getTextFrame().getTextRange().getParagraphs().get(0).getRuns().add("ume");
        run2.getFont().getColor().setRGB(Color.GetYellow());
        ITextRange run3 = rectangle.getTextFrame().getTextRange().getParagraphs().get(0).getRuns().add("nts");
        run3.getFont().getColor().setRGB(Color.GetLightBlue());
        ITextRange paragraph0 = rectangle.getTextFrame().getTextRange().getParagraphs().get(0);
        paragraph0.getFont().setSize(36);
        paragraph0.getFont().setBold(true);

        rectangle.getTextFrame().getTextRange().getParagraphs().add(" ");

        // Add second paragraph
        ITextRange paragraph1 = rectangle.getTextFrame().getTextRange().getParagraphs().add();
        ITextRange run4 = paragraph1.getRuns().add("Fast, efficient Excel, Word, and PDF APIs for .NET Standard 2.0 and Java");
        run4.getFont().getColor().setRGB(Color.GetBlack());
        run4.getFont().setSize(20);
        run4.getFont().setName("Agency FB");

        rectangle.getTextFrame().getTextRange().getParagraphs().add(" ");

        // Add third paragraph
        ITextRange paragraph2 = rectangle.getTextFrame().getTextRange().getParagraphs().add();
        ITextRange run5 = paragraph2.getRuns().add("Take total document control with ultra-fast, low-footprint document APIs for enterprise apps");
        run5.getFont().getColor().setRGB(Color.GetBlack());
        run5.getFont().setSize(16);
        run5.getFont().setName("Times New Roman");
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

}
