package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MoveSignatureLines extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet activeSheet = workbook.getActiveSheet();
        ISignatureSet signatures = workbook.getSignatures();

        // Prepare data
        ISignature signatureShinzo = signatures.addSignatureLine(activeSheet, 100.0, 50.0);
        ISignatureSetup setup = signatureShinzo.getSetup();
        setup.setShowSignDate(false);
        setup.setAllowComments(false);
        setup.setSigningInstructions("Please check the content before signing.");
        setup.setSuggestedSigner("Shinzo Nagama");
        setup.setSuggestedSignerEmail("shinzo.nagama@ea.com");
        setup.setSuggestedSignerLine2("Commander (Balanced)");

        // Move signature line
        IShape signatureLineShape = signatureShinzo.getSignatureLineShape();
        signatureLineShape.setTop(signatureLineShape.getTop() + 100);
        signatureLineShape.setLeft(signatureLineShape.getLeft() + 50);
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
