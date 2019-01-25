package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.OpenFileFormat;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportCsvFileToWorkbook extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/Information.csv"), OpenFileFormat.Csv);
    }

    @Override
    public String getTemplateName() {
        return "Information.csv";
    }

    @Override
    public String[] getResources() {
        return new String[]{"xlsx/Information.csv"};
    }
}
