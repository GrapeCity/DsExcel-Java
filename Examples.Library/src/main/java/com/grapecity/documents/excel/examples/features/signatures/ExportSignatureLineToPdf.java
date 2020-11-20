package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ExportSignatureLineToPdf extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        workbook.open(getResourceStream("xlsx/Signature.xlsx"));
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] {"xlsx/Signature.xlsx"};
    }

    @Override
    public String[] getDependencies() {
        return new String[]{ "compile group: 'com.grapecity.documents', name: 'gcexcel.extension', version: '3.2.0'" };
    }
}
