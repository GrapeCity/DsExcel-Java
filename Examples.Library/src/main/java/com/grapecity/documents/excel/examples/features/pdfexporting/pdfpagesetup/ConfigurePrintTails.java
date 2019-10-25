package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigurePrintTails extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/RepeatTails.xlsx"));
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Repeat the columns "I" at the left of each page when saving worksheet2 to pdf.
        IWorksheet worksheet1 = workbook.getWorksheets().get(0);
        worksheet1.getPageSetup().setPrintTailColumns("$I:$I");

        //Repeat the row 50th at the bottom of each page when saving worksheet1 to pdf.
        IWorksheet worksheet2 = workbook.getWorksheets().get(1);
        worksheet2.getPageSetup().setPrintTailRows("$50:$50");
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
    
	@Override
	public String getTemplateName() {
		return "RepeatTails.xlsx";
	}

	@Override
	public String[] getResources() {
		return new String[] { "xlsx/RepeatTails.xlsx" };
	}
}
