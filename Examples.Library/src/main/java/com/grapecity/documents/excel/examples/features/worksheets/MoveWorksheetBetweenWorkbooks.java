package com.grapecity.documents.excel.examples.features.worksheets;

import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MoveWorksheetBetweenWorkbooks extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file Home inventory.xlsx from resource to the source workbook
        Workbook source_workbook = new Workbook();
        InputStream source_fileStream = this.getResourceStream("xlsx/Home inventory.xlsx");
        source_workbook.open(source_fileStream);

        //Move content of active sheet from source workbook to the current workbook before the first sheet
        IWorksheet move_worksheet = source_workbook.getActiveSheet().moveBefore(workbook.getWorksheets().get(0));
        move_worksheet.setName("Move of Home inventory");
        move_worksheet.activate();
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