package com.grapecity.documents.excel.examples.features.formulas.precedentsanddependents;

import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class GetAllPrecedents extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        worksheet.getRange("E2").setFormula("=sum(C1:C2)");
        worksheet.getRange("C1").setFormula("=B1");
        worksheet.getRange("B1").setFormula("=sum(A1:A2)");
        worksheet.getRange("A1").setValue(1);
        worksheet.getRange("A2").setValue(2);
        worksheet.getRange("C2").setValue(3);

        ArrayList<IRange> list = new ArrayList<IRange>();
        for (IRange item : worksheet.getRange("E2").getPrecedents())
        {
        	list.add(item);
        }

        while (list.size() > 0)
        {
        	ArrayList<IRange> temp = list;
            list = new ArrayList<IRange>();
            for (IRange item : temp)
            {
                for (int i = 0; i < item.getRowCount(); i++)
                {
                    for (int j = 0; j < item.getColumnCount(); j++)
                    {
                    	List<IRange> dependents = item.getCells().get(i, j).getPrecedents();
                        if (dependents.size() == 0)
                        {
                            item.getCells().get(i, j).getInterior().setColor(Color.GetSkyBlue());
                        }
                        else
                        {
                        	item.getCells().get(i, j).getInterior().setColor(Color.GetGray());
                            list.addAll(dependents);
                        }
                    }
                }
            }
        }
    }
}
