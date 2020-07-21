package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ToJsonFromJsonForWorksheet extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        //ToJson&FromJson can be used in combination with spread.sheets product:http://spread.grapecity.com/spreadjs/sheets/

        //GrapeCity Documents for Excel import an excel file.
        //Change the path to real source file path.
        String source = "source.xlsx";
        workbook.open(source);

        //Open the same user file
        Workbook new_workbook = new Workbook();
        new_workbook.open(source);

        for (IWorksheet worksheet : workbook.getWorksheets()) {
        	//Do any change of worksheet
        	//...

        	//GrapeCity Documents for Excel export a worksheet to a json string.
        	String json = worksheet.toJson();
        	//Use the json string to initialize spread.sheets product.
        	//Product spread.sheets will show the excel file contents.

        	//Use spread.sheets product export a json string of worksheet.
        	//GrapeCity Documents for Excel use the json string to update content of the corresponding worksheet.
        	new_workbook.getWorksheets().get(worksheet.getName()).fromJson(json);
        }

        //GrapeCity Documents for Excel export workbook to an excel file.
        //Change the path to real export file path.
        String export = "export.xlsx";
        new_workbook.save(export);
        
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
