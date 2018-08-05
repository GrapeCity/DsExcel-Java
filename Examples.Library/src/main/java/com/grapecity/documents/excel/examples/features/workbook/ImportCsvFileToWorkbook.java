package com.grapecity.documents.excel.examples.features.workbook;

import java.net.URL;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportCsvFileToWorkbook extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        URL url = ClassLoader.getSystemResource("xlsx/Information.csv");
        //Open csv file stream.
        String filePath = url.getFile();
        //TODO OPEN .csv uncompleted
        workbook.open(filePath);
    }

}
