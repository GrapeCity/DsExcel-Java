package com.grapecity.documents.excel.examples.features.signatures;

import java.io.InputStream;

import com.grapecity.documents.excel.ISignature;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.XlsxOpenOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

public class DeleteSignatureFromSignatureLine extends ExampleBase {
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

        // Remove signature of the first signed signature line.
        for (ISignature signature : workbook.getSignatures()) {
            if (signature.getIsSignatureLine() && signature.getIsSigned()) {
                // Remove digital signature.
                // The signature line will not be removed from the SignatureSet
                // in digital signature only mode.
                // Because signature lines are shapes.
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
        return new String[] { "xlsx/SignatureLine2Signed1NotSigned.xlsx" };
    }
}