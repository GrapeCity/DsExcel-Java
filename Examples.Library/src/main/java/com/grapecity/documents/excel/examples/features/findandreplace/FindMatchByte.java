package com.grapecity.documents.excel.examples.features.findandreplace;

import java.util.Locale;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FindOptions;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FindMatchByte extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
        // This option is valid when culture is ja-JP or zh-CN.
        workbook.setCulture(Locale.JAPANESE);
        ;
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        // Prepare data
        worksheet.getRange("A1:A4")
                .setValue(new String[] { "Mario Games", "スーパーマリオブラザーズ", "ﾏﾘｵ&ﾙｲｰｼﾞRPG3 DX", "マリオ＆ルイージRPG1 DX" });
        
        // Find the first cell that contains "マリオ" (match width)
        // and mark it with red foreground.
        IRange searchRange = worksheet.getUsedRange();
        FindOptions matchByteOptions = new FindOptions();
        matchByteOptions.setMatchByte(true);
        
        IRange marioCell = searchRange.find("マリオ", null, matchByteOptions);
        marioCell.getFont().setColor(Color.GetRed());
        
        // Find the first cell that contains "ルイージ" (ignore width)
        // and mark it with green background.
        IRange luigiCell = searchRange.find("ルイージ");
        luigiCell.getInterior().setColor(Color.GetGreen());

	}
}