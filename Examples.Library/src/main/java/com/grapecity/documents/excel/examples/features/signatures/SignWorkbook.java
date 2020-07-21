package com.grapecity.documents.excel.examples.features.signatures;

import java.io.IOException;
import java.io.InputStream;
import java.security.KeyStore;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.security.cert.CertificateException;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SignWorkbook extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        ISignature signature = workbook.getSignatures().addNonVisibleSignature();

        SignatureDetails details = new SignatureDetails();
        details.setAddress1("<your address>");
        details.setAddress2("<address 2>");
        details.setSignatureComments("Final");
        details.setCity("<your city>");
        details.setStateOrProvince("<your state or province>");
        details.setPostalCode("<your postal code>");
        details.setCountryName("<your country or region>");
        details.setClaimedRole("<your role>");
        details.setCommitmentTypeDescription("Approved");
        details.setCommitmentTypeQualifier("Final");

        try {
            KeyStore ks = KeyStore.getInstance("pkcs12");
            String password = "<your password>";
            char[] passwordChars = password.toCharArray();
            String pfxFileKey = "<your certificate file>";
            InputStream pfxStrm = getClass().getClassLoader().getResourceAsStream(pfxFileKey);
            try {
                ks.load(pfxStrm, passwordChars);
            } catch (NoSuchAlgorithmException e) {
                throw new RuntimeException(e);
            } catch (CertificateException e) {
                throw new RuntimeException(e);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            signature.sign(ks, password, details);
            // Commit signature
            workbook.save("SignWorkbook.xlsx");
        } catch (KeyStoreException e) {
            throw new RuntimeException(e);
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

    @Override
    public boolean getShowScreenshot() {
        return true;
    }

    @Override
    public boolean getCanDownload() {
        return false;
    }
}
