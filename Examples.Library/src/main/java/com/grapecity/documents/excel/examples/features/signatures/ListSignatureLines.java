package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ListSignatureLines extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet activeSheet = workbook.getActiveSheet();
        ISignatureSet signatures = workbook.getSignatures();

        // Prepare data
        ISignature signatureShinzo = signatures.addSignatureLine(activeSheet, 100.0, 50.0);
        ISignatureSetup setup1 = signatureShinzo.getSetup();
        setup1.setShowSignDate(false);
        setup1.setAllowComments(false);
        setup1.setSigningInstructions("Please check the content before signing.");
        setup1.setSuggestedSigner("Shinzo Nagama");
        setup1.setSuggestedSignerEmail("shinzo.nagama@ea.com");
        setup1.setSuggestedSignerLine2("Commander (Balanced)");

        ISignature signatureKenji = signatures.addSignatureLine(activeSheet, 100.0, 350.0);
        ISignatureSetup setup2 = signatureKenji.getSetup();
        setup2.setShowSignDate(true);
        setup2.setAllowComments(true);
        setup2.setSigningInstructions("Please check the content before signing!");
        setup2.setSuggestedSigner("Kenji Tenzai");
        setup2.setSuggestedSignerEmail("kenji.tenzai@ea.com");
        setup2.setSuggestedSignerLine2("Commander (Mecha)");

        // List signatures with indexes
        for (int i = 0; i < signatures.getCount(); i++) {
            ISignature signature = signatures.get(i);
            // Insert your code here

        }

        for (ISignature signature : signatures) {
            // Insert your code here

        }
    }

    @Override
    public boolean getIsNew() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}
