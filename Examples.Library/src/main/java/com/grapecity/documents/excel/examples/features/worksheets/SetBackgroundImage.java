package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class SetBackgroundImage extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Use sheet index to get worksheet.
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        InputStream inputStream = this.getResourceStream("logo.png");
        try {
            byte[] bytes = new byte[inputStream.available()];
            inputStream.read(bytes, 0, bytes.length);
            worksheet.setBackgroundPicture(bytes);
        }catch (IOException ioe){
            ioe.printStackTrace();
        }
    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public boolean getIsNew() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[]{"logo.png"};
    }
}
