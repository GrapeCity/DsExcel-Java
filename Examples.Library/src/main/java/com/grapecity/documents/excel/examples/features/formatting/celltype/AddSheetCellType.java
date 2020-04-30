package com.grapecity.documents.excel.examples.features.formatting.celltype;

import com.grapecity.documents.excel.ButtonCellType;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddSheetCellType extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        ButtonCellType buttonCellType = new ButtonCellType();
        buttonCellType.setText("Button");
        buttonCellType.setButtonBackColor("Azure");
        buttonCellType.setMarginLeft(10);
        buttonCellType.setMarginRight(10);

        worksheet.setCellType(buttonCellType);
    }
}
