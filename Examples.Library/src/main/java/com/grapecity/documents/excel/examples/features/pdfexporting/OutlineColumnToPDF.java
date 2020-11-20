package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ImageSource;
import com.grapecity.documents.excel.SummaryRow;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class OutlineColumnToPDF extends ExampleBase {
    @Override
    public void execute(Workbook workbook){
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set data.
        Object[][] data = new Object[][]{
            { "Preface",                                   "1",      1,  0 } ,
            { "Java SE5 and SE6",                          "1.1",    2,  1 },
            { "Java SE6",                                  "1.1.1",  2,  2 },
            { "The 4th edition",                           "1.2",    2,  1 },
            { "Changes",                                   "1.2.1",  3,  2 },
            { "Note on the cover design",                  "1.3",    4,  1 },
            { "Acknowledgements",                          "1.4",    4,  1 },
            { "Introduction",                              "2",      9,  0 },
            { "Prerequisites",                             "2.1",    9,  1 },
            { "Learning Java",                             "2.2",    10, 1 },
            { "Goals",                                     "2.3",    10, 1 },
            { "Teaching from this book",                   "2.4",    11, 1 },
            { "JDK HTML documentation",                    "2.5",    11, 1 },
            { "Exercises",                                 "2.6",    12, 1 },
            { "Foundations for Java",                      "2.7",    12, 1 },
            { "Source code",                               "2.8",    12, 1 },
            { "Coding standards",                          "2.8.1",  14, 2 },
            { "Errors",                                    "2.9",    14, 1 },
            { "Introduction to Objects",                   "3",      15, 0 },
            { "The progress of abstraction",               "3.1",    15, 1 },
            { "An object has an interface",                "3.2",    17, 1 },
            { "An object provides services",               "3.3",    18, 1 },
            { "The hidden implementation",                 "3.4",    19, 1 },
            { "Reusing the implementation",                "3.5",    20, 1 },
            { "Inheritance",                               "3.6",    21, 1 },
            { "Is-a vs. is-like-a relationships",          "3.6.1",  24, 2 },
            { "Interchangeable objects with polymorphism", "3.7",    25, 1 },
            { "The singly rooted hierarchy",               "3.8",    28, 1 },
            { "Containers",                                "3.9",    28, 1 },
            { "Parameterized types (Generics)",            "3.10",   29, 1 },
            { "Object creation & lifetime",                "3.11",   30, 1 },
            { "Exception handling: dealing with errors",   "3.12",   31, 1 },
            { "Concurrent programming",                    "3.13",   32, 1 },
            { "Java and the Internet",                     "3.14",   33, 1 },
            { "What is the Web?",                          "3.14.1", 33, 2 },
            { "Client-side programming",                   "3.14.2", 34, 2 },
            { "Server-side programming",                   "3.14.3", 38, 2 },
            { "Summary",                                   "3.15",   38, 1 }
        };
        worksheet.getRange("A1:C38").setValue(data);

        //Set ColumnWidth.
        worksheet.getRange("A:A").setColumnWidthInPixel(310);
        worksheet.getRange("B:C").setColumnWidthInPixel(150);

        //Set IndentLevel.
        for (int i = 0; i < data.length; i++)
        {
            worksheet.getRange(i, 0).setIndentLevel((int)data[i][3]);
        }

        //Show the summary row above the detail rows.
        worksheet.getOutline().setSummaryRow(SummaryRow.Above);

        //Don't show the row outline when interacting with SJS, the exported excel file still show the row outline.
        worksheet.setShowRowOutline(false);

        //Set outline column, the corresponding row outlines will also be automatically created.
        worksheet.getOutlineColumn().setColumnIndex(0);
        worksheet.getOutlineColumn().setShowCheckBox(true);
        worksheet.getOutlineColumn().setShowImage(true);
        worksheet.getOutlineColumn().setMaxLevel(2);
        worksheet.getOutlineColumn().getImages().add(new ImageSource(this.getResourceStream("archiverFolder.png"), ImageType.PNG));
        worksheet.getOutlineColumn().getImages().add(new ImageSource(this.getResourceStream("newFloder.png"), ImageType.PNG));
        worksheet.getOutlineColumn().getImages().add(new ImageSource(this.getResourceStream("docFile.png"), ImageType.PNG));
        worksheet.getOutlineColumn().setCollapseIndicator(new ImageSource(this.getResourceStream("decreaseIndicator.png"), ImageType.PNG));
        worksheet.getOutlineColumn().setExpandIndicator(new ImageSource(this.getResourceStream("increaseIndicator.png"), ImageType.PNG));
        worksheet.getOutlineColumn().setCheckStatus(0, true);
        worksheet.getOutlineColumn().setCollapsed(1, true);
        
        //Print the headings&gridlines.
        worksheet.getPageSetup().setPrintHeadings(true);
        worksheet.getPageSetup().setPrintGridlines(true);
	}

	@Override
	public boolean getSavePdf() {
        return true;
	}

	@Override
	public boolean getShowViewer() {
        return false;
	}

	@Override
	public String[] getResources() {
        return new String[] { "archiverFolder.png", "newFloder.png", "docFile.png", "increaseIndicator.png", "decreaseIndicator.png"  };
	}
}

