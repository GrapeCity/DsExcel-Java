package com.grapecity.documents.excel.examples.features.signatures;

import java.io.InputStream;

import com.grapecity.documents.excel.ISignature;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.XlsxOpenOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

public class DeleteNonVisibleSignature extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        // This file contains 2 already signed signature lines,
        // a not signed signature line, and a non-visible signature.
        InputStream signedFile = getResourceStream("xlsx/MixedDigitalSignatures.xlsx");

        // Use DigitalSignatureOnly mode, because the workbook was already signed.
        // If you don't open it with digital signature only mode,
        // all existing signatures will be removed after saving the workbook.
        XlsxOpenOptions openOption = new XlsxOpenOptions();
        openOption.setDigitalSignatureOnly(true);
        workbook.open(signedFile, openOption);

        // Remove the first non-visible signature.
        for (ISignature signature : workbook.getSignatures()) {
            if (!signature.getIsSignatureLine() && signature.getIsSigned()) {
                // Remove digital signature.
                // The signature will be removed from the SignatureSet.
                signature.delete();
                break;
            }
        }

        // Commit signatures
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
    public String[] getResources() {
        return new String[] { "xlsx/MixedDigitalSignatures.xlsx" };
    }
}