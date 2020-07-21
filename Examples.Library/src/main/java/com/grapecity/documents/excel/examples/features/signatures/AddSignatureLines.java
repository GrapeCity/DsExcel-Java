package com.grapecity.documents.excel.examples.features.signatures;

import java.io.IOException;
import java.io.InputStream;
import java.security.KeyStore;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.security.cert.CertificateException;

import com.grapecity.documents.excel.*;
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
