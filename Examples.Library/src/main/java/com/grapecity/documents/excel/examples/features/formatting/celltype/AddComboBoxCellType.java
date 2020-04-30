package com.grapecity.documents.excel.examples.features.formatting.celltype;

import com.grapecity.documents.excel.ComboBoxCellItem;
import com.grapecity.documents.excel.ComboBoxCellType;
import com.grapecity.documents.excel.EditorValueType;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddComboBoxCellType extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        ComboBoxCellType cellType = new ComboBoxCellType();
        cellType.setEditorValueType(EditorValueType.Value);
        ComboBoxCellItem item = new ComboBoxCellItem();
        item.setValue("US");
        item.setText("United States");
        cellType.getItems().add(item);
        item = new ComboBoxCellItem();
        item.setValue("CN");
        item.setText("China");
        cellType.getItems().add(item);
        item = new ComboBoxCellItem();
        item.setValue("JP");
        item.setText("Japan");
        cellType.getItems().add(item);

        worksheet.getRange("C5").setCellType(cellType);
        worksheet.getRange("C5").setValue("CN");
    }
}
