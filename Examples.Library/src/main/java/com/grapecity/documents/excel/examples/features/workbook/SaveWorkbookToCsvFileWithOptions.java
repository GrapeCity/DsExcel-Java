package com.grapecity.documents.excel.examples.features.workbook;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.CsvSaveOptions;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SaveWorkbookToCsvFileWithOptions extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        Object data = new Object[][]{
            {"Name", "City", "Birthday", "Sex", "Weight", "Height"},
            {"Bob", "NewYork", new GregorianCalendar(1968, 5, 8), "male", 80, 180},
            {"Betty", "NewYork", new GregorianCalendar(1972, 6, 3), "female", 72, 168},
            {"Gary", "NewYork", new GregorianCalendar(1964, 2, 2), "male", 71, 179},
            {"Hunk", "Washington", new GregorianCalendar(1972, 7, 8), "male", 80, 171},
            {"Cherry", "Washington", new GregorianCalendar(1986, 1, 2), "female", 58, 161},
            {"Eva", "Washington", new GregorianCalendar(1993, 2, 5), "female", 71, 180}
        };

        // Set data.
        IWorksheet sheet = workbook.getWorksheets().get(0);
        sheet.getRange("A1:F7").setValue(data);
        sheet.getTables().add(sheet.getRange("A1:F7"), true);

        //Save csv options
        CsvSaveOptions options = new CsvSaveOptions();
        options.setColumnSeparator("-");

        //Change the path to real export path when save.
        workbook.save("dest.csv", options);

    }

	@Override
	public boolean getCanDownload() {
		return false;
	}

    @Override
    public boolean getShowViewer() {

        return false;

    }
}
