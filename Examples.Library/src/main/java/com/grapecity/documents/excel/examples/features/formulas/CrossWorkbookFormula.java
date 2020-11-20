package com.grapecity.documents.excel.examples.features.formulas;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CrossWorkbookFormula extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.getWorksheets().get(0).getRange("B1").setFormula("='[SourceWorkbook.xlsx]Sheet1'!A1");

        // create a new workbook as the instance of external workbook.
        Workbook workbook2 = new Workbook();
        workbook2.getWorksheets().get(0).getRange("A1").setValue("Hello, World!");
        // update the caches of external workbook data.
        for (String item : workbook.getExcelLinkSources())
        {
            workbook.updateExcelLink(item, workbook2);
        }
    }
    @Override
    public boolean getIsNew()
    {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}