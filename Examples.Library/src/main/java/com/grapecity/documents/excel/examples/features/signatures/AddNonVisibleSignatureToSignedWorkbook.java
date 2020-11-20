package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.ISignature;
import com.grapecity.documents.excel.SignatureDetails;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.XlsxOpenOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.IOException;
import java.io.InputStream;
import java.security.KeyStore;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.security.cert.CertificateException;

public class AddNonVisibleSignatureToSignedWorkbook extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        // This file contains 2 already signed signature lines and
        // a not signed signature line.
        InputStream signedFile = getResourceStream("xlsx/SignatureLine2Signed1NotSigned.xlsx");

        // Use DigitalSignatureOnly mode, because the workbook was already signed.
        // If you don't open it with digital signature only mode,
        // all existing signatures will be removed after saving the workbook.
        XlsxOpenOptions openOption = new XlsxOpenOptions();
        openOption.setDigitalSignatureOnly(true);
        workbook.open(signedFile, openOption);

        // TODO: Use your certificate.
        KeyStore ks;
        try {
            ks = KeyStore.getInstance("pkcs12");
        } catch (KeyStoreException e) {
            throw new RuntimeException(e);
        }

        // TODO: Use your password.
        String password = "<your password>";
        char[] passwordChars = password.toCharArray();
        String pfxFileKey = "<your certificate file>";
        InputStream pfxStrm = getResourceStream(pfxFileKey);
        try {
            ks.load(pfxStrm, passwordChars);
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException(e);
        } catch (CertificateException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        // TODO: Fill details with your information.
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

        // Add signature to this workbook.
        ISignature signature = workbook.getSignatures().addNonVisibleSignature();
        signature.sign(ks, password, details);

        // Commit signatures
        workbook.save("AddNonVisibleSignatureToSignedWorkbook.xlsx");
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public boolean getCanDownload() {
        return false;
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/SignatureLine2Signed1NotSigned.xlsx" };
    }

    @Override
    public String[] getDependencies() {
        return new String[]{ "compile group: 'com.grapecity.documents', name: 'gcexcel.extension', version: '3.2.0'" };
    }

    @Override
    public String[] getImportPackages() {
        return new String[]{
                "import java.io.IOException;",
                "import java.io.InputStream;",
                "import java.security.KeyStore;",
                "import java.security.KeyStoreException;",
                "import java.security.NoSuchAlgorithmException;",
                "import java.security.cert.CertificateException;"
        };
    }
}