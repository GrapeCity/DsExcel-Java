package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

import java.io.InputStream;
import java.util.stream.Stream;

public class ConvertExcelToPDF extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        //Open an excel file
        InputStream fileStream = this.getTemplateStream();
        workbook.open(fileStream);
    }

    @Override
    public String getTemplateName() {
        return "Employee absence schedule.xlsx";
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

}