package com.grapecity.documents.excel.examples.features.hyperlinks;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateShapeWithHyperlink extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Add shapes
        IShape shape1 = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 10, 0, 100, 100);
        shape1.getTextFrame().getTextRange().add("Go to google web site.");
        IShape shape2 = worksheet.getShapes().addShape(AutoShapeType.RightArrow, 10, 120, 100, 100);
        shape2.getTextFrame().getTextRange().add("Go to sheet1 C3:E4");
        IShape shape3 = worksheet.getShapes().addShape(AutoShapeType.Oval, 10, 240, 100, 100);
        shape3.getTextFrame().getTextRange().add("Send an email to sales");
        IShape shape4 = worksheet.getShapes().addShape(AutoShapeType.LeftArrow, 10, 360, 100, 100);
        shape4.getTextFrame().getTextRange().add("Link to external.xlsx file.");

        //add a hyperlink link to web page.
        worksheet.getHyperlinks().add(shape1, "http://www.google.com/", null, "open google web site.", "Google");

        //add a hyperlink link to a range in this document.
        worksheet.getHyperlinks().add(shape2, null, "Sheet1!$C$3:$E$4", "Go to sheet1 C3:E4", null);

        //add a hyperlink link to email address.
        worksheet.getHyperlinks().add(shape3, "mailto:us.sales@grapecity.com", null, "Send an email to sales", "Send an email to sales");

        //add a hyperlink link to external file.
        //change the path to real picture file path.
        String path = "external.xlsx";
        worksheet.getHyperlinks().add(shape4, path, null, "link to external.xlsx file.", "External.xlsx");
    }
    
    @Override
    public boolean getShowViewer() {
        return false;
    }
}
