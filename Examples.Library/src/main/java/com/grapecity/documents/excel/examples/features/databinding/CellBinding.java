package com.grapecity.documents.excel.examples.features.databinding;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CellBinding extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        //#region Define custom classes
        //public class SalesRecord
        //{
        //    public string Area;
        //    public string Salesman;
        //    public string Product;
        //    public string ProductType;
        //    public int Sales;
        //}
        //#endregion

        //#region Init data
        SalesRecord record = new SalesRecord();
        record.area = "NorthChina";
        record.salesman = "Hellen";
        record.product = "Apple";
        record.productType = "Fruit";
        record.sales = 120;
        //#endregion

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Set binding path for cell.
        worksheet.getRange("A1").setBindingPath("area");
        worksheet.getRange("B2").setBindingPath("salesman");
        worksheet.getRange("C2").setBindingPath("product");
        worksheet.getRange("D3").setBindingPath("productType");

        // Set data source.
        worksheet.setDataSource(record);
    }

    @Override
    public boolean getIsNew() {
        return true;
    }

    @Override
    public String[] getRefs() {
        return new String[]{"SalesRecord"};
    }
}
