package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ExportSignatureLineToPdf extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        workbook.open(getResourceStream("xlsx/Signature.xlsx"));		
    }

    @Override
    public boolean getIsNew() {
        return true;
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
}
