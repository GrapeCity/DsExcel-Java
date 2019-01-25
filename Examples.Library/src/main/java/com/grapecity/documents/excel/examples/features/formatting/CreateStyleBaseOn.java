package com.grapecity.documents.excel.examples.features.formatting;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.IStyle;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Pattern;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.UnderlineType;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.Tutorial;

public class CreateStyleBaseOn extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").setStyle(workbook.getStyles().get("Good"));
        worksheet.getRange("A1").setValue("Good");

        // Create and modify a style based on current existing style
        IStyle myGood = workbook.getStyles().add("MyGood", workbook.getStyles().get("Good"));
        myGood.getFont().setBold(true);
        myGood.getFont().setItalic(true);

        worksheet.getRange("B1").setStyle(workbook.getStyles().get("MyGood"));
        worksheet.getRange("B1").setValue("MyGood");
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
