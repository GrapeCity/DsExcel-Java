package com.grapecity.documents.excel.examples.features.rangeoperations;

import java.io.InputStream;
import java.util.EnumSet;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.PasteType;
import com.grapecity.documents.excel.UsedRangeType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CutCopyRangeBetweenWorkbooks extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file Home inventory.xlsx from resource
        Workbook source_workbook = new Workbook();
        InputStream source_fileStream = this.getResourceStream("xlsx/Home inventory.xlsx");
        source_workbook.open(source_fileStream);

        //Hide gridline
        workbook.getActiveSheet().getSheetView().setDisplayGridlines(false);

        workbook.getActiveSheet().getRange("A1").setValue("Copy content from the first sheet of source workbook");
        workbook.getActiveSheet().getRange("A1").getFont().setColor(Color.GetRed());
        workbook.getActiveSheet().getRange("A1").getFont().setBold(true);

        //Copy content of active sheet from source workbook to the current sheet at A2
        source_workbook.getActiveSheet().getUsedRange(EnumSet.allOf(UsedRangeType.class)).copy(workbook.getActiveSheet().getRange("A2"),
                EnumSet.of(PasteType.Formulas, PasteType.Formats, PasteType.RowHeights, PasteType.ColumnWidths));

        workbook.getActiveSheet().getRange("C21").setValue("Cut content from the second sheet of source workbook");
        workbook.getActiveSheet().getRange("C21").getFont().setColor(Color.GetRed());
        workbook.getActiveSheet().getRange("C21").getFont().setBold(true);

        //Cut content of second sheet from source workbook to the current sheet at C22
        source_workbook.getWorksheets().get(1).getRange("2:15").cut(workbook.getActiveSheet().getRange("C22"));
        
        //Make the theme of two workbooks same
        workbook.setTheme(source_workbook.getTheme());
	}


	@Override
	public String getTemplateName() {
        return "Home inventory.xlsx";
	}

	@Override
	public String[] getResources() {
        return new String[] { "xlsx/Home inventory.xlsx" };
	}
}