package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CopySignatureLines extends ExampleBase {
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

        // Copy signature line with Range.Copy
        IRange srcRange = activeSheet.getRange("A1:I15");
        IRange destRange = activeSheet.getRange("A16:I30");
        srcRange.copy(destRange);

        // Copy signature line with Shape.Duplicate
        signature.getSignatureLineShape().duplicate();

        // Copy signature line with Worksheet.Copy
        activeSheet.copy();
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
    public boolean getShowScreenshot() {
        return true;
    }
}
