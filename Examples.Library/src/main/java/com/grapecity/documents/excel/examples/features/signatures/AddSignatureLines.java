package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.ISignature;
import com.grapecity.documents.excel.ISignatureSetup;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddSignatureLines extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        ISignature signature = workbook.getSignatures().addSignatureLine(
            workbook.getActiveSheet(), 100.0, 50.0);

        ISignatureSetup setup = signature.getSetup();
        setup.setShowSignDate(true);
        setup.setAllowComments(true);
        setup.setSigningInstructions("<your signing instructions>");
        setup.setSuggestedSigner("<signer's name>");
        setup.setSuggestedSignerEmail("example@microsoft.com");
        setup.setSuggestedSignerLine2("<signer's title>");
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
