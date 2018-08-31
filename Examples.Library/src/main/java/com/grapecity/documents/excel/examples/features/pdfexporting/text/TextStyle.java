package com.grapecity.documents.excel.examples.features.pdfexporting.text;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class TextStyle extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

//C# TO JAVA CONVERTER TODO TASK: There is no preprocessor in Java:
        ///#region Aligment
        sheet.getRange("A1").setValue("Alignment");

        sheet.getRange("B2").setValue("Left Alignment");
        sheet.getRange("B2").setHorizontalAlignment(HorizontalAlignment.Left);

        sheet.getRange("C2").setValue("Center Alignment");
        sheet.getRange("C2").setHorizontalAlignment(HorizontalAlignment.Center);

        sheet.getRange("D2").setValue("Right Alignment");
        sheet.getRange("D2").setHorizontalAlignment(HorizontalAlignment.Right);

        sheet.getRange("B3").setValue("Top Alignment");
        sheet.getRange("B3").setVerticalAlignment(VerticalAlignment.Top);

        sheet.getRange("C3").setValue("Middle Alignment");
        sheet.getRange("C3").setVerticalAlignment(VerticalAlignment.Center);

        sheet.getRange("D3").setValue("Bottom Alignment");
        sheet.getRange("D3").setVerticalAlignment(VerticalAlignment.Bottom);

        sheet.getRange("B4").setValue("This is  a test string for Justify Alignment. \nThis is a test string for Justify Alignment. ");
        sheet.getRange("B4").setHorizontalAlignment(HorizontalAlignment.Justify);
        sheet.getRange("B4").setVerticalAlignment(VerticalAlignment.Justify);

        sheet.getRange("C4").setValue("This is  a test string for Distributed Alignment. \nThis is a test string for Distributed Alignment. ");
        sheet.getRange("C4").setHorizontalAlignment(HorizontalAlignment.Distributed);
        sheet.getRange("C4").setVerticalAlignment(VerticalAlignment.Distributed);

        ///#endregion

        //Wordwrap
        sheet.getRange("A6").setValue("Wordwrap");
        sheet.getRange("B7").setValue("This is a test string for Wordwrap");
        sheet.getRange("C7").setValue("This is a test string \n for Wordwrap");
        sheet.getRange("B7:C7").setWrapText(true);

        //Indent
        sheet.getRange("A9").setValue("Indent");
        sheet.getRange("B10").setValue("Left Indent");
        sheet.getRange("B10").setIndentLevel(3);
        sheet.getRange("C10").setValue("Right Indent");
        sheet.getRange("C10").setIndentLevel(3);
        sheet.getRange("C10").setHorizontalAlignment(HorizontalAlignment.Right);

        //Shrink to fit
        sheet.getRange("A12").setValue("Shrink to fit");
        sheet.getRange("B13").setValue("This is a test string for \"Shrink to fit\"");
        sheet.getRange("B13").setShrinkToFit(true);

        //Underline
        sheet.getRange("A15").setValue("Underline");
        sheet.getRange("B16").setValue("Single Underline");
        sheet.getRange("B16").getFont().setUnderline(UnderlineType.Single);

        //Strikthrough
        sheet.getRange("A18").setValue("Strikthrough");
        sheet.getRange("B19").setValue("Strikthrough");
        sheet.getRange("B19").getFont().setStrikethrough(true);

        //Other settings
        sheet.getColumns().get(0).getFont().setBold(true);
        sheet.getColumns().get(0).setColumnWidthInPixel(100);
        sheet.getColumns().get(1).setColumnWidthInPixel(200);
        sheet.getColumns().get(2).setColumnWidthInPixel(245);
        sheet.getColumns().get(3).setColumnWidthInPixel(234);
        sheet.getRows().get(2).setRowHeightInPixel(72);
        sheet.getRows().get(3).setRowHeightInPixel(123);
        sheet.getRows().get(6).setRowHeightInPixel(48);

        sheet.getRange("A1:D19").getBorders().setLineStyle(BorderLineStyle.Thin);
        sheet.getPageSetup().setPaperSize(PaperSize.A3);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}