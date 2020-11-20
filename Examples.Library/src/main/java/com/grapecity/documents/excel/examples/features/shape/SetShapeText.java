package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetShapeText extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.Parallelogram, 1, 1, 200, 100);
        shape.setWidth(500);
        shape.setHeight(200);

        shape.getTextFrame().getTextRange().getFont().getColor().setRGB(Color.FromArgb(0, 255, 0));
        shape.getTextFrame().getTextRange().getFont().setBold(true);
        shape.getTextFrame().getTextRange().getFont().setItalic(true);
        shape.getTextFrame().getTextRange().getFont().setSize(20);
        shape.getTextFrame().getTextRange().getFont().setStrikethrough(true);

        shape.getTextFrame().getTextRange().getParagraphs().add("This is a parallelogram shape.");
        shape.getTextFrame().getTextRange().getParagraphs().add("My name is XXX");
        shape.getTextFrame().getTextRange().getParagraphs().get(1).getRuns().add("Hello World!");

        shape.getTextFrame().getTextRange().getParagraphs().get(1).getRuns().get(0).getFont().setStrikethrough(false);
        shape.getTextFrame().getTextRange().getParagraphs().get(1).getRuns().get(0).getFont().setSize(35);

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
