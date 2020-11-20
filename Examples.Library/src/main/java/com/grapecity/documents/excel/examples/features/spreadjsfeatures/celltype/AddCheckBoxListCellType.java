package com.grapecity.documents.excel.examples.features.spreadjsfeatures.celltype;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddCheckBoxListCellType extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        CheckBoxListCellType cellType = new CheckBoxListCellType();
        cellType.setDirection(CellTypeDirection.Horizontal);
        cellType.setTextAlign(CellTypeTextAlign.Right);
        cellType.setIsFlowLayout(false);
        cellType.setMaxColumnCount(2);
        cellType.setMaxRowCount(1);
        cellType.setHorizontalSpacing(20);
        cellType.setVerticalSpacing(5);

        cellType.getItems().add(new SelectFieldItem("sample1", "1"));
        cellType.getItems().add(new SelectFieldItem("sample2", "2"));
        cellType.getItems().add(new SelectFieldItem("sample3", "3"));
        cellType.getItems().add(new SelectFieldItem("sample4", "4"));
        cellType.getItems().add(new SelectFieldItem("sample5", "5"));

        worksheet.getRange("A1").setRowHeight(60);
        worksheet.getRange("A1").setColumnWidth(25);

        worksheet.getRange("A1").setCellType(cellType);
    }

    @Override
    public boolean getIsNew(){
        return true;
    }

    @Override
    public boolean getSavePdf()
    {
        return true;
    }
}

