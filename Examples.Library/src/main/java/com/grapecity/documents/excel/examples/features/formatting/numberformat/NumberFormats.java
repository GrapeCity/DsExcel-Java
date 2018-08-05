package com.grapecity.documents.excel.examples.features.formatting.numberformat;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class NumberFormats extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A:H").setColumnWidth(17);

        // Display 111 as 111.
        worksheet.getRange("A1").setValue(111);
        worksheet.getRange("A1").setNumberFormat("#####");

        // Display 222 as 00222.
        worksheet.getRange("B1").setValue(222);
        worksheet.getRange("B1").setNumberFormat("00000");

        // Display 12345678 as 12,345,678.
        worksheet.getRange("C1").setValue(12345678);
        worksheet.getRange("C1").setNumberFormat("#,#");

        // Display .126 as 0.13.
        worksheet.getRange("D1").setValue(.126);
        worksheet.getRange("D1").setNumberFormat("0.##");

        // Display 74.4 as 74.400.
        worksheet.getRange("E1").setValue(74.4);
        worksheet.getRange("E1").setNumberFormat("##.000");

        // Display 1.6 as 160.0%.
        worksheet.getRange("F1").setValue(1.6);
        worksheet.getRange("F1").setNumberFormat("0.0%");

        // Display 4321 as $4,321.00.
        worksheet.getRange("G1").setValue(4321);
        worksheet.getRange("G1").setNumberFormat("$#,##0.00");

        // Display 8.75 as 8 3/4.
        worksheet.getRange("H1").setValue(8.75);
        worksheet.getRange("H1").setNumberFormat("# ?/?");
    }

}
