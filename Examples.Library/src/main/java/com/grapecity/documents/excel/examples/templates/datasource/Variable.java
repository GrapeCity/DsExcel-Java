package com.grapecity.documents.excel.examples.templates.datasource;

import java.io.InputStream;
import java.util.ArrayList;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Variable extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file Template_StudentInfo.xlsx from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_StudentInfo.xlsx");
        workbook.open(templateFile);

        ///#region Define custom classes
        //        public class StudentInfo
        //        {
        //        	public String name;
        //        	public String address;
        //        	public ArrayList<Family> family;
        //        }

        ///#region Init Data
        ArrayList<StudentInfo> studentInfos = new ArrayList<StudentInfo>();

        StudentInfo student1 = new StudentInfo();
        student1.name = "Jane";
        student1.address = "101, Halford Avenue, Fremont, CA";
        studentInfos.add(student1);

        StudentInfo student2 = new StudentInfo();
        student2.name = "Mark";
        student2.address = "2005 Klamath Ave APT, Santa Clara, CA";
        studentInfos.add(student2);

        StudentInfo student3 = new StudentInfo();
        student3.name = "Carol";
        student3.address = "1063 E EI Camino Real, Sunnyvale, CA 94087, USA";
        studentInfos.add(student3);

        StudentInfo student4 = new StudentInfo();
        student4.name = "Liano";
        student4.address = "1977 St Lawrence Dr, Santa Clara, CA 95051, USA";
        studentInfos.add(student4);

        StudentInfo student5 = new StudentInfo();
        student5.name = "Hellen";
        student5.address = "3661 Peacock Ct, Santa Clara, CA 95051, USA";
        studentInfos.add(student5);

        String className = "Class 3";
        ///#endregion

        //Add data source
        workbook.addDataSource("className", className);
        workbook.addDataSource("s", studentInfos);

        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_StudentInfo.xlsx";
	}

	@Override
	public boolean getHasTemplate() {
        return true;
	}
	
	@Override
	public boolean getShowTemplate() {
        return true;
    }

	@Override
	public String[] getResources() {
        return new String[] { "xlsx/Template_StudentInfo.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
        return new String[] { "StudentInfo", "Family", "Guardian" };
	}
}