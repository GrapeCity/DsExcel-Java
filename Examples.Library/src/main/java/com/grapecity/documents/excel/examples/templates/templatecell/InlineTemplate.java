package com.grapecity.documents.excel.examples.templates.templatecell;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class InlineTemplate extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
		//Load template file Template_Score.xlsx from resource
		InputStream templateFile = this.getResourceStream("xlsx/Template_InlineTemplate.xlsx");
		workbook.open(templateFile);

		/// #region Define custom classes
		//		public class DataInfo {
		//			public int id;
		//			public String name;
		//			public int score;
		//			public String team;
		//		}
		/// #endregion

		///#region Init Data
		List<DataInfo> datasource = new ArrayList<DataInfo>();

		DataInfo r1 = new DataInfo();
		r1.id = 10;
		r1.name = "Bob";
		r1.score = 12;
		r1.team = "Xi'An";
		datasource.add(r1);

		DataInfo r2 = new DataInfo();
		r2.id = 11;
		r2.name = "Tommy";
		r2.score = 6;
		r2.team = "Xi'An";
		datasource.add(r2);

		DataInfo r3 = new DataInfo();
		r3.id = 12;
		r3.name = "Jaguar";
		r3.score = 15;
		r3.team = "Xi'An";
		datasource.add(r3);

		DataInfo r4 = new DataInfo();
		r4.id = 2;
		r4.name = "Phillip";
		r4.score = 9;
		r4.team = "BeiJing";
		datasource.add(r4);

		DataInfo r5 = new DataInfo();
		r5.id = 3;
		r5.name = "Hunter";
		r5.score = 10;
		r5.team = "BeiJing";
		datasource.add(r5);

		DataInfo r6 = new DataInfo();
		r6.id = 4;
		r6.name = "Hellen";
		r6.score = 8;
		r6.team = "BeiJing";
		datasource.add(r6);

		DataInfo r7 = new DataInfo();
		r7.id = 5;
		r7.name = "Jim";
		r7.score = 9;
		r7.team = "BeiJing";
		datasource.add(r7);
		///#endregion

		//Add data source
		workbook.addDataSource("ds", datasource);
		//Invoke to process the template
		workbook.processTemplate();
	}

	@Override
	public boolean getIsNew() {
		return true;
	}

	@Override
	public String getTemplateName() {
		return "Template_InlineTemplate.xlsx";
	}

	@Override
	public boolean getHasTemplate() {
		return true;
	}

	@Override
	public String[] getResources() {
		return new String[] { "xlsx/Template_InlineTemplate.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
		return new String[] { "DataInfo" };
	}
}