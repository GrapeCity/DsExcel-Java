package com.grapecity.documents.excel.examples.features.databinding;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CellBinding extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        //#region Define custom classes
        //public class SalesRecord {
        //    public String area;
        //    public String city;
        //    public String category;
        //    public String name;
        //    public double revenue;
        //}
        //#endregion

        //#region Init data
        SalesRecord record = new SalesRecord();
        record.area = "North America";
        record.city = "Chicago";
        record.category = "Consumer Electronics";
        record.name = "Bose 785593-0050";
        record.revenue = 92800;
        //#endregion

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Set binding path for cell.
        worksheet.getRange("A1").setBindingPath("area");
        worksheet.getRange("B2").setBindingPath("city");
        worksheet.getRange("C2").setBindingPath("name");
        worksheet.getRange("D3").setBindingPath("revenue");

        // Set data source.
        worksheet.setDataSource(record);
    }

    @Override
    public String[] getRefs() {
        return new String[]{"SalesRecord"};
    }
}
