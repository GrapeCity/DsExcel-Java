package com.grapecity.documents.excel.examples.features.pagesetup;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigPrintTailColumns extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/RepeatTails.xlsx"));
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        // Repeat the column "I" at the right of each page when saving pdf.
        worksheet.getPageSetup().setPrintTailColumns("$I:$I");

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
