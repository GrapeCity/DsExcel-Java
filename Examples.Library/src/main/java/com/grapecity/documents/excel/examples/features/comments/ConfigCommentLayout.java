package com.grapecity.documents.excel.examples.features.comments;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IComment;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.LineDashStyle;
import com.grapecity.documents.excel.drawing.LineStyle;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigCommentLayout extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IComment commentC3 = worksheet.getRange("C3").addComment("Range C3's comment.");
        commentC3.getShape().getLine().getColor().setRGB(Color.GetLightGreen());
        commentC3.getShape().getLine().setWeight(3);
        commentC3.getShape().getLine().setStyle(LineStyle.ThickThin);
        commentC3.getShape().getLine().setDashStyle(LineDashStyle.Solid);
        commentC3.getShape().getFill().getColor().setRGB(Color.GetPink());
        commentC3.getShape().setWidth(100);
        commentC3.getShape().setHeight(200);
        commentC3.getShape().getTextFrame().getTextRange().getFont().setBold(true);
        commentC3.setVisible(true);

    }

}
