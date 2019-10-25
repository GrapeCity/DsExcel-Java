package com.grapecity.documents.excel.examples.templates.datasource;

import java.io.InputStream;
import java.io.InputStreamReader;

import com.google.gson.Gson;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class JsonFile extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//Load template file Template_FamilyInfo.xlsx from resource
		InputStream templateFile = this.getResourceStream("xlsx/Template_FamilyInfo.xlsx");
		workbook.open(templateFile);

		///#region Define custom classes
		//		public class StudentInfos
		//		{
		//			public ArrayList<StudentInfo> student;
		//		}

		//		public class StudentInfo
		//		{
		//			public String name;
		//			public String address;
		//			public ArrayList<Family> family;
		//		}

		//		public class Family
		//		{
		//			public Guardian father;
		//			public Guardian mother;
		//		}

		//		public class Guardian
		//		{
		//			public String name;
		//			public String occupation;
		//		}
		///#endregion

		//Get data from json file
		InputStreamReader reader = new InputStreamReader(this.getResourceStream("Template_FamilyInfo.json"));
		Gson gson = new Gson();
		StudentInfos datasource = gson.fromJson(reader, StudentInfos.class);

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
		return "Template_FamilyInfo.xlsx";
	}

	@Override
	public boolean getHasTemplate() {
		return true;
	}

	@Override
	public String[] getResources() {
		return new String[] { "xlsx/Template_FamilyInfo.xlsx", "Template_FamilyInfo.json" };
	}
	
	@Override
	public String[] getRefs() {
		return new String[] { "StudentInfos", "StudentInfo", "Family", "Guardian" };
	}
}