package com.grapecity.documents.excel.examples.features.formatting.celltype;

import com.grapecity.documents.excel.CheckBoxAlign;
import com.grapecity.documents.excel.CheckBoxCellType;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddCheckBoxCellType extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        CheckBoxCellType cellType = new CheckBoxCellType();
        cellType.setCaption("Caption");
        cellType.setTextTrue("True");
        cellType.setTextFalse("False");
        cellType.setTextIndeterminate("Indeterminate");
        cellType.setIsThreeState(true);
        cellType.setTextAlign(CheckBoxAlign.Right);

        worksheet.getRange("C5:C6").setCellType(cellType);
        worksheet.getRange("C5").setValue(true);
        worksheet.getRange("C6").setValue(false);
    }
}
