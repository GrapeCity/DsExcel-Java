package com.grapecity.documents.excel.examples.features.pagesetup;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigPrintTailRows extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/RepeatTails.xlsx"));
        IWorksheet worksheet = workbook.getWorksheets().get(1);
        
        //Repeat the row 50th at the bottom of each page when saving pdf.
        worksheet.getPageSetup().setPrintTailRows("$50:$50");

	}

	@Override
	public String getTemplateName() {
		return "RepeatTails.xlsx";
	}

	@Override
	public String[] getResources() {
		return new String[] { "xlsx/RepeatTails.xlsx" };
	}

	@Override
	public boolean getShowViewer() {
		return false;
	}

}
