package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.OpenFileFormat;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportCsvFileToWorkbook extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getTemplateStream(), OpenFileFormat.Csv);
    }

    @Override
    public String getTemplateName() {
        return "Information.csv";
    }
}
