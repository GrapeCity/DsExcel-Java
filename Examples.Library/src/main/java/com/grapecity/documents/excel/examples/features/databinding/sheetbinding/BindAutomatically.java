package com.grapecity.documents.excel.examples.features.databinding.sheetbinding;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.features.databinding.SalesData;
import com.grapecity.documents.excel.examples.features.databinding.SalesRecord;

import java.util.ArrayList;

public class BindAutomatically extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        //#region Define custom classes
        //public class SalesData
        //{
        //    public ArrayList<SalesRecord> Records;
        //}

        //public class SalesRecord
        //{
        //    public String Area;
        //    public String Salesman;
        //    public String Product;
        //    public String ProductType;
        //    public int Sales;
        //}
        //#endregion

        //#region Init data
        SalesData datasource = new SalesData();
        datasource.records = new ArrayList<SalesRecord>();

        SalesRecord record1 = new SalesRecord();
        record1.area = "NorthChina";
        record1.salesman = "Hellen";
        record1.product = "Apple";
        record1.productType = "Fruit";
        record1.sales = 120;
        datasource.records.add(record1);

        SalesRecord record2 = new SalesRecord();
        record2.area = "NorthChina";
        record2.salesman = "Hellen";
        record2.product = "Banana";
        record2.productType = "Fruit";
        record2.sales = 143;
        datasource.records.add(record2);

        SalesRecord record3 = new SalesRecord();
        record3.area = "NorthChina";
        record3.salesman = "Hellen";
        record3.product = "Kiwi";
        record3.productType = "Fruit";
        record3.sales = 322;
        datasource.records.add(record3);
        //#endregion

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Data source would be bound to columns automatically.
        worksheet.setDataSource(datasource.records);
    }

    @Override
    public String[] getRefs() {
        return new String[]{"SalesData", "SalesRecord"};
    }
}

