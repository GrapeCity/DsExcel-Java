package com.grapecity.documents.excel.examples.features.comments;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IComment;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetRichTextForComment extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IComment commentC3 = worksheet.getRange("C3").addComment("This is a rich text comment:\r\n");

        //config the paragraph's style.
        commentC3.getShape().getTextFrame().getTextRange().getParagraphs().get(0).getFont().setBold(true);

        //add runs for the paragraph.
        commentC3.getShape().getTextFrame().getTextRange().getParagraphs().get(0).getRuns().add("Run1 font size is 15.", 1);
        commentC3.getShape().getTextFrame().getTextRange().getParagraphs().get(0).getRuns().add("Run2 font strikethrough.", 2);
        commentC3.getShape().getTextFrame().getTextRange().getParagraphs().get(0).getRuns().add("Run3 font italic, green color.");

        //config the first run of the paragraph's style.
        commentC3.getShape().getTextFrame().getTextRange().getParagraphs().get(0).getRuns().get(1).getFont().setSize(15);
        //config the second run of the paragraph's style. 
        commentC3.getShape().getTextFrame().getTextRange().getParagraphs().get(0).getRuns().get(2).getFont().setStrikethrough(true);

        //config the third run of the paragraph's style. 
        commentC3.getShape().getTextFrame().getTextRange().getParagraphs().get(0).getRuns().get(3).getFont().setItalic(true);
        commentC3.getShape().getTextFrame().getTextRange().getParagraphs().get(0).getRuns().get(3).getFont().getColor().setRGB(Color.GetGreen());

        //show comment.
        commentC3.setVisible(true);

        commentC3.getShape().setWidthInPixel(300);
        commentC3.getShape().setHeightInPixel(100);

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
