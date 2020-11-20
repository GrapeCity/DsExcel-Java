package com.grapecity.documents.excel.examples.features.spreadjsfeatures.celltype;

import com.grapecity.documents.excel.ButtonCellType;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddButtonCellType extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        ButtonCellType cellType = new ButtonCellType();
        cellType.setText("Hello");
        cellType.setButtonBackColor("Azure");
        cellType.setMarginLeft(10);
        cellType.setMarginRight(10);

        worksheet.getRange("C5").setCellType(cellType);
    }

    @Override
    public boolean getSavePdf()
    {
        return true;
    }
}
