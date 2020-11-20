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
        //public class SalesData {
        //    public ArrayList<SalesRecord> sales;
        //}

        //public class SalesRecord {
        //    public String area;
        //    public String city;
        //    public String category;
        //    public String name;
        //    public double revenue;
        //}
        //#endregion

        //#region Init data
        SalesData datasource = new SalesData();
        datasource.sales = new ArrayList<SalesRecord>();

        SalesRecord record1 = new SalesRecord();
        record1.area = "North America";
        record1.city = "Chicago";
        record1.category = "Consumer Electronics";
        record1.name = "Bose 785593-0050";
        record1.revenue = 92800;
        datasource.sales.add(record1);

        SalesRecord record2 = new SalesRecord();
        record2.area = "North America";
        record2.city = "New York";
        record2.category = "Consumer Electronics";
        record2.name = "Bose 785593-0050";
        record2.revenue = 92800;
        datasource.sales.add(record2);

        SalesRecord record3 = new SalesRecord();
        record3.area = "South America";
        record3.city = "Santiago";
        record3.category = "Consumer Electronics";
        record3.name = "Bose 785593-0050";
        record3.revenue = 19550;
        datasource.sales.add(record3);
        //#endregion

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Data source would be bound to columns automatically.
        worksheet.setDataSource(datasource.sales);
    }

    @Override
    public String[] getRefs() {
        return new String[]{"SalesData", "SalesRecord"};
    }
}

