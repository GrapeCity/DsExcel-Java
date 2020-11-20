package com.grapecity.documents.excel.examples.features.signatures;

import com.grapecity.documents.excel.ISignature;
import com.grapecity.documents.excel.ISignatureSet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.security.cert.CertificateExpiredException;
import java.security.cert.CertificateNotYetValidException;
import java.security.cert.X509Certificate;

public class VerifySignatures extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        workbook.open(getResourceStream("xlsx/Signature.xlsx"));

        ISignatureSet signatures = workbook.getSignatures();

        boolean signed = false;
        boolean valid = false;
        X509Certificate certificate = null;

        // Verify the first signature
        for (ISignature signature : signatures) {
            if (signature.getIsSigned()) {
                // Save the result in locals. You can print them later.
                signed = true;
                certificate = signature.getDetails().getSignatureCertificate();
                valid = signature.getIsValid();
                break;
            }
        }

        // Verify the first certificate

        boolean certificateIsValid = true;

        // Check expiration date and start date
        try {
            certificate.checkValidity();
        } catch (CertificateExpiredException e) {
            certificateIsValid = false;
            return;
        } catch (CertificateNotYetValidException e) {
            certificateIsValid = false;
            return;
        }

        // TODO: Verify the certificate chain

        // Important:
        // We don't know how to get certificate chain on this platform.
        // Certificate verification needs to get certificate chain from additional places, 
        // because the embedded certificate doesn't include a complete certificate chain.
        // On .NET platform, we can use
        // System.Security.Cryptography.X509Certificates.X509Chain
        // to get the complete certificate chain from Windows certificate storage or
        // the OpenSSL based PowerShell certificate storage.
        // There's no equivalent of X509Chain on this platform.
        // The implementation of X509Chain can be found here:
        // https://github.com/dotnet/runtime/tree/master/src/libraries/System.Security.Cryptography.X509Certificates/src/Internal/Cryptography
        // It supports Windows, Linux and Mac.
        // But we don't think it's convertible to Java.
        // Because it has pointer types, nullable reference types, 
        // ByRef-like value types and unsigned integer types.

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
    public String[] getDependencies() {
        return new String[]{ "compile group: 'com.grapecity.documents', name: 'gcexcel.extension', version: '3.2.0'" };
    }

    @Override
    public String[] getImportPackages() {
        return new String[]{
                "import java.security.cert.CertificateExpiredException;",
                "import java.security.cert.CertificateNotYetValidException;",
                "import java.security.cert.X509Certificate;"
        };
    }


    @Override
    public String[] getResources() {
        return new String[] {"xlsx/Signature.xlsx"};
    }
}
