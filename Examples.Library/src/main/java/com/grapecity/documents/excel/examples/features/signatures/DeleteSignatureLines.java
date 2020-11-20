package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.ISignature;
import com.grapecity.documents.excel.ISignatureSetup;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.util.concurrent.Callable;

public class DeleteSignatureLines extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet activeSheet = workbook.getActiveSheet();

        // Create a anonymous class for creating new signature line with predefined data
        Callable<ISignature> newSignatureLine = new Callable<ISignature>() {
            @Override
            public ISignature call() throws Exception {
                ISignature signature = workbook.getSignatures().addSignatureLine(activeSheet, 100.0, 50.0);
                ISignatureSetup setup = signature.getSetup();
                setup.setShowSignDate(false);
                setup.setAllowComments(false);
                setup.setSigningInstructions("Please check the content before signing.");
                setup.setSuggestedSigner("Shinzo Nagama");
                setup.setSuggestedSignerEmail("shinzo.nagama@ea.com");
                setup.setSuggestedSignerLine2("Commander (Balanced)");
                return signature;
            }
        };

        // Create a new signature line and delete with Signature.Delete
        ISignature signatureForTest;
        try {
            signatureForTest = newSignatureLine.call();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        signatureForTest.delete();

        // Create a new signature line and delete with Shape.Delete
        try {
            signatureForTest = newSignatureLine.call();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        IShape signatureLineShape = signatureForTest.getSignatureLineShape();
        signatureLineShape.delete();
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public String[] getDependencies() {
        return new String[]{ "compile group: 'com.grapecity.documents', name: 'gcexcel.extension', version: '3.2.0'" };
    }

    @Override
    public String[] getImportPackages() {
        return new String[]{"import java.util.concurrent.Callable;"};
    }
}
