package com.grapecity.documents.excel.examples.features.spreadjsfeatures.celltype;

import com.grapecity.documents.excel.ButtonCellType;
import com.grapecity.documents.excel.CheckBoxAlign;
import com.grapecity.documents.excel.CheckBoxCellType;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddRowColumnCellType extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getColumns().get(3).setColumnWidthInPixel(100);
        CheckBoxCellType cellType = new CheckBoxCellType();
        cellType.setCaption("CheckBox");
        cellType.setTextTrue("True");
        cellType.setTextFalse("False");
        cellType.setIsThreeState(true);
        cellType.setTextAlign(CheckBoxAlign.Right);
        worksheet.getColumns().get(3).setCellType(cellType);
        worksheet.getRange("D1:D10").setValue(true);

        ButtonCellType buttonCellType = new ButtonCellType();
        buttonCellType.setText("Button");
        buttonCellType.setButtonBackColor("Azure");
        buttonCellType.setMarginLeft(10);
        buttonCellType.setMarginRight(10);

        worksheet.getRows().get(3).setCellType(buttonCellType);
    }

    @Override
    public boolean getSavePdf()
    {
        return true;
    }
}
