package com.grapecity.documents.excel.examples.features.formatting.celltype;

import com.grapecity.documents.excel.HyperLinkCellType;
import com.grapecity.documents.excel.HyperLinkTargetType;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddHyperlinkCellType extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        HyperLinkCellType cellType = new HyperLinkCellType();
        cellType.setText("Google");
        cellType.setLinkColor("Blue");
        cellType.setLinkToolTip("Search by google");
        cellType.setVisitedLinkColor("Green");
        cellType.setTarget(HyperLinkTargetType.Blank);

        worksheet.getRange("C5").setCellType(cellType);
        worksheet.getRange("C5").setValue("http://www.google.com");
    }

	@Override
	public boolean getIsNew() {
        return true;
    }
}
