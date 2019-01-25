package com.grapecity.documents.excel.examples.exceltemplates;

import java.net.URL;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FinancialKPIs extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Load template file Financial KPIs.xlsx from resource
        workbook.open(this.getResourceStream("xlsx/Financial KPIs.xlsx"));

        IWorksheet worksheet = workbook.getActiveSheet();

        //set values
        Object proData = new Object[][]{
                {1483550, 121386},
                {0.4336, 0.32},
                {0.1236, -0.0438},
                {12.36, -0.0438}
        };

        Object proData1 = new Object[]{0.14, 0.0029, 0.0361, 0.0361};

        worksheet.getRange("E7:F10").setValue(proData);
        worksheet.getRange("I7:I10").setValue(proData1);

        Object activeData = new Object[][]{
                {"4.32 item", "2.00 item"},
                {"72 days", "45 days"},
                {"53 days", "55 days"}
        };

        Object activeData1 = new Object[]{"0.45 items", "-5 days", "-6 days"};

        worksheet.getRange("E12:F14").setValue(activeData);
        worksheet.getRange("I12:I14").setValue(activeData1);

        Object effData = new Object[][]{
                {0.3446, 0.25},
                {0.5335, 0.11}
        };

        Object effData1 = new Object[]{0.1245, 0.1946};

        worksheet.getRange("E16:F17").setValue(effData);
        worksheet.getRange("I16:I17").setValue(effData1);

        Object linData = new Object[][]{
                {"0.91:1", "'2:1"},
                {"0.58:1", "'1:1"}
        };

        Object linData1 = new Object[]{"'0.02:1", "'0.03:1"};

        worksheet.getRange("E19:F20").setValue(linData);
        worksheet.getRange("I19:I20").setValue(linData1);

        Object geaData = new Object[][]{
                {-9.60, 0.85},
                {0.68, 0.5}
        };

        Object geaData1 = new Object[]{6.65, 0.0282};

        worksheet.getRange("E22:F23").setValue(geaData);
        worksheet.getRange("I22:I23").setValue(geaData1);

        Object casData = new Object[][]{
                {0.0735, 1.2},
                {0.1442, 0.1442}
        };

        Object casData1 = new Object[]{-0.0046, 0.023};

        worksheet.getRange("E25:F26").setValue(casData);
        worksheet.getRange("I25:I26").setValue(casData1);

    }


    @Override
    public String getTemplateName() {

        return "Financial KPIs.xlsx";

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public String[] getResources() {
        return new String[] {"xlsx/Financial KPIs.xlsx"};
    }
}
