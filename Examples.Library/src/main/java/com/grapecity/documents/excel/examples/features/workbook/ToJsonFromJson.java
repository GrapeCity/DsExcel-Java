package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ToJsonFromJson extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        // /ToJson&FromJson can be used in combination with spread.sheets product:http://spread.grapecity.com/spreadjs/sheets/

        //GrapeCity Documents for Excel import an excel file.
        //change the path to real source file path.
        String source = "source.xlsx";
        workbook.open(source);
        //GrapeCity Documents for Excel export to a json string.
        String jsonstr = workbook.toJson();
        //use the json string to initialize spread.sheets product.
        //spread.sheets will show the excel file contents.

        //spread.sheets product export a json string.
        //GrapeCity Documents for Excel use the json string to initialize.
        workbook.fromJson(jsonstr);
        //GrapeCity Documents for Excel export workbook to an excel file.
        //change the path to real export file path.
        String export = "export.xlsx";
        workbook.save(export);
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
