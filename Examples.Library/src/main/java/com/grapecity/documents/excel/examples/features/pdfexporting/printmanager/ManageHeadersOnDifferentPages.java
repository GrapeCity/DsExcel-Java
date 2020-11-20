package com.grapecity.documents.excel.examples.features.pdfexporting.printmanager;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PageInfo;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.RepeatSetting;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ManageHeadersOnDifferentPages extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        workbook.open(this.getResourceStream("xlsx/MultipleHeaders.xlsx"));

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        List<RepeatSetting> repeatSettings = new ArrayList<RepeatSetting>();
        
        //The title rows of the "B2:F87" is "$2:$2"
        RepeatSetting repeatSetting = new RepeatSetting();
        repeatSetting.setTitleRowStart(1);
        repeatSetting.setTitleRowEnd(1);
        repeatSetting.setRange(worksheet.getRange("B2:F87"));
        repeatSettings.add(repeatSetting);

        //The title rows of the "B91:F189" is "$91:$91"
        RepeatSetting repeatSetting2 = new RepeatSetting();
        repeatSetting2.setTitleRowStart(88);
        repeatSetting2.setTitleRowEnd(88);
        repeatSetting2.setRange(worksheet.getRange("B89:F149"));
        repeatSettings.add(repeatSetting2);

        //Create a PrintManager.
        PrintManager printManager = new PrintManager();

        //Get the pagination information of the worksheet.
        List<PageInfo> pages = printManager.paginate(worksheet, null, repeatSettings);

        //Save the modified pages into pdf file.
        printManager.savePageInfosToPDF(outputStream, pages);
	}

	@Override
	public boolean getSavePageInfos() {
        return true;
	}

	@Override
	public boolean getShowViewer() {
        return false;
	}

	@Override
	public String getTemplateName() {
        return "MultipleHeaders.xlsx";
	}

	@Override
	public String[] getResources() {
        return new String[] { "xlsx/MultipleHeaders.xlsx" };
	}

}
