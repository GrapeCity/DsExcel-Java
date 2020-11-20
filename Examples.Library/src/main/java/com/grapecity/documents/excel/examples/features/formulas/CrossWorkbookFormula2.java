package com.grapecity.documents.excel.examples.features.formulas;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CrossWorkbookFormula2 extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/Urban Centers 2017.xlsx"));
        Workbook flints = new Workbook();
        flints.open(this.getResourceStream("xlsx/Flints.xlsx"));
        Workbook jackson = new Workbook();
        jackson.open(this.getResourceStream("xlsx/Jackson.xlsx"));
        Workbook petrosky = new Workbook();
        petrosky.open(this.getResourceStream("xlsx/Petrosky.xlsx"));

        workbook.getWorksheets().get(0).getRange("B6:E11").setFormula("='[Flints.xlsx]Sheet1'!B7+'[Jackson.xlsx]Sheet1'!B6+'[Petrosky.xlsx]Sheet1'!B7");
        workbook.updateExcelLink("Flints.xlsx", flints);
        workbook.updateExcelLink("Jackson.xlsx", jackson);
        workbook.updateExcelLink("Petrosky.xlsx", petrosky);
        workbook.calculate();
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/Urban Centers 2017.xlsx", "xlsx/Flints.xlsx", "xlsx/Jackson.xlsx", "xlsx/Petrosky.xlsx"};
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
