package com.grapecity.documents.excel.examples.templates.templatesamples;

import java.io.InputStream;
import java.util.ArrayList;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class DepartmentBudget extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_DepartmentBudget.xlsx");
        workbook.open(templateFile);

        //#region Define custom classes
        //class Departments
        //{
        //    public List<Department> dpt;
        //}

        //class Department
        //{
        //    public string name;
        //    public string mgr;
        //    public double bud;
        //    public List<Employee> emp;
        //}

        //class Employee
        //{
        //    public string name;
        //    public double salary;
        //}
        //#endregion

        //#region Init Data
        Departments departments = new Departments();
        departments.dpt = new ArrayList<Department>();

        //Department 1
        Department department1 = new Department();
        department1.name = "Marketing";
        department1.mgr = "Carl Sommerset";
        department1.bud = 354586;

        department1.emp = new ArrayList<Employee>();

        Employee d1_e1 = new Employee();
        d1_e1.name = "JoeKline";
        d1_e1.salary = 49402;
        department1.emp.add(d1_e1);

        Employee d1_e2 = new Employee();
        d1_e2.name = "Lisa Crane";
        d1_e2.salary = 81337;
        department1.emp.add(d1_e2);

        Employee d1_e3 = new Employee();
        d1_e3.name = "John Ryes";
        d1_e3.salary = 43503;
        department1.emp.add(d1_e3);

        Employee d1_e4 = new Employee();
        d1_e4.name = "Elli Davidson";
        d1_e4.salary = 67334;
        department1.emp.add(d1_e4);

        Employee d1_e5 = new Employee();
        d1_e5.name = "Jack Reaze";
        d1_e5.salary = 68314;
        department1.emp.add(d1_e5);

        Employee d1_e6 = new Employee();
        d1_e6.name = "Ben Lam";
        d1_e6.salary = 44696;
        department1.emp.add(d1_e6);

        departments.dpt.add(department1);

        //Department 2
        Department department2 = new Department();
        department2.name = "Sales";
        department2.mgr = "Kelly Johnson";
        department2.bud = 237721;

        department2.emp = new ArrayList<Employee>();
        Employee d2_e1 = new Employee();
        d2_e1.name = "Liam Elmerson";
        d2_e1.salary = 61892;
        department2.emp.add(d2_e1);

        Employee d2_e2 = new Employee();
        d2_e2.name = "Angela Sanderson";
        d2_e2.salary = 38020;
        department2.emp.add(d2_e2);

        Employee d2_e3 = new Employee();
        d2_e3.name = "Blake Schwarz";
        d2_e3.salary = 55701;
        department2.emp.add(d2_e3);

        Employee d2_e4 = new Employee();
        d2_e4.name = "Linda Barataz";
        d2_e4.salary = 82108;
        department2.emp.add(d2_e4);

        departments.dpt.add(department2);

        //Department 3
        Department department3 = new Department();
        department3.name = "Engineering";
        department3.mgr = "Gina Davis";
        department3.bud = 624789;

        department3.emp = new ArrayList<Employee>();

        Employee d3_e1 = new Employee();
        d3_e1.name = "Christopher Dean";
        d3_e1.salary = 58329;
        department3.emp.add(d3_e1);

        Employee d3_e2 = new Employee();
        d3_e2.name = "Jack Linner";
        d3_e2.salary = 63684;
        department3.emp.add(d3_e2);

        Employee d3_e3 = new Employee();
        d3_e3.name = "Cathy Raines";
        d3_e3.salary = 73147;
        department3.emp.add(d3_e3);

        Employee d3_e4 = new Employee();
        d3_e4.name = "Scott Ashton";
        d3_e4.salary = 77213;
        department3.emp.add(d3_e4);

        Employee d3_e5 = new Employee();
        d3_e5.name = "Larry Wisell";
        d3_e5.salary = 72796;
        department3.emp.add(d3_e5);

        Employee d3_e6 = new Employee();
        d3_e6.name = "Bart Ingram";
        d3_e6.salary = 50009;
        department3.emp.add(d3_e6);

        Employee d3_e7 = new Employee();
        d3_e7.name = "Wesley Page";
        d3_e7.salary = 82378;
        department3.emp.add(d3_e7);

        Employee d3_e8 = new Employee();
        d3_e8.name = "Alan Keyes";
        d3_e8.salary = 67105;
        department3.emp.add(d3_e8);

        Employee d3_e9 = new Employee();
        d3_e9.name = "Wilson Musk";
        d3_e9.salary = 80128;
        department3.emp.add(d3_e9);

        departments.dpt.add(department3);

        //Department 4
        Department department4 = new Department();
        department4.name = "Customer Service";
        department4.mgr = "Kenneth Smith";
        department4.bud = 127596;

        department4.emp = new ArrayList<Employee>();

        Employee d4_e1 = new Employee();
        d4_e1.name = "Sherry Meeks";
        d4_e1.salary = 38919;
        department4.emp.add(d4_e1);

        Employee d4_e2 = new Employee();
        d4_e2.name = "Sharon Reeves";
        d4_e2.salary = 40963;
        department4.emp.add(d4_e2);

        Employee d4_e3 = new Employee();
        d4_e3.name = "Max Devillo";
        d4_e3.salary = 47714;
        department4.emp.add(d4_e3);

        departments.dpt.add(department4);
        //         #endregion

        //Add data source
        workbook.addDataSource("ds", departments);
        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_DepartmentBudget.xlsx";
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
        return new String[] { "xlsx/Template_DepartmentBudget.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
        return new String[] { "Departments", "Department", "Employee" };
	}
}