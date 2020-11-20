package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.ISignature;
import com.grapecity.documents.excel.ISignatureSetup;
import com.grapecity.documents.excel.SignatureDetails;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.KeyStore;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.security.cert.CertificateException;

public class CountersignSignature extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        ISignature signature = workbook.getSignatures().addSignatureLine(workbook.getActiveSheet(), 100.0, 50.0);

        ISignatureSetup setup = signature.getSetup();
        setup.setShowSignDate(true);
        setup.setAllowComments(true);
        setup.setSigningInstructions("<your signing instructions>");
        setup.setSuggestedSigner("<signer's name>");
        setup.setSuggestedSignerEmail("example@microsoft.com");
        setup.setSuggestedSignerLine2("<signer's title>");

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
            signature.sign(ks, password, "<your name>", details);

            // Reopen the workbook to commit signature
            ByteArrayOutputStream outStrm = new ByteArrayOutputStream();
            workbook.save(outStrm);
            ByteArrayInputStream inStrm = new ByteArrayInputStream(outStrm.toByteArray());
            workbook.open(inStrm);

            // Modify
            workbook.getWorksheets().get("Sheet1").getRange("A1").setValue("Modified");

            // Countersign
            workbook.getSignatures().get(0).countersign(ks, password);

            // Commit signature
            workbook.save("CountersignSignature.xlsx");
        } catch (KeyStoreException e) {
            throw new RuntimeException(e);
        }
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

    @Override
    public String[] getResources() {
        return new String[]{""};
    }
}
