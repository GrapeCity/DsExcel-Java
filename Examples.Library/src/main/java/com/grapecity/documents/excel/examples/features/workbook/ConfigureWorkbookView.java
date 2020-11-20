package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.IWorkbookView;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigureWorkbookView extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        // Workbook view settings.
        IWorkbookView bookView = workbook.getBookView();
        bookView.setDisplayVerticalScrollBar(false);
        bookView.setDisplayWorkbookTabs(true);
        bookView.setTabRatio(0.5);

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
