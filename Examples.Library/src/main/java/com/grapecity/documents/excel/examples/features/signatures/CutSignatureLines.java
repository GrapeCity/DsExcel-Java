package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CutSignatureLines extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet activeSheet = workbook.getActiveSheet();
        ISignature signature = workbook.getSignatures().addSignatureLine(activeSheet, 100.0, 50.0);

        ISignatureSetup setup = signature.getSetup();
        setup.setShowSignDate(false);
        setup.setAllowComments(false);
        setup.setSigningInstructions("Please check the content before signing.");
        setup.setSuggestedSigner("Shinzo Nagama");
        setup.setSuggestedSignerEmail("shinzo.nagama@ea.com");
        setup.setSuggestedSignerLine2("Commander (Balanced)");

        // Cut signature line with Range.Cut
        IRange srcRange = activeSheet.getRange("A1:I15");
        IRange destRange = activeSheet.getRange("A16:I30");
        srcRange.cut(destRange);
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public boolean getShowScreenshot() {
        return true;
    }

    @Override
    public String[] getDependencies() {
        return new String[]{ "compile group: 'com.grapecity.documents', name: 'gcexcel.extension', version: '3.2.0'" };
    }
}
