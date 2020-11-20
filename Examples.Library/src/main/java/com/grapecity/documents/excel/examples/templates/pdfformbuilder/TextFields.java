package com.grapecity.documents.excel.examples.templates.pdfformbuilder;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class TextFields extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		// Load template file from resource
		InputStream templateFile = this.getResourceStream("xlsx/Template_TextFields.xlsx");
		workbook.open(templateFile);

		// #region Init Data
		List<AddressBook> addressBooks = new ArrayList<AddressBook>();

		AddressBook addressBook1 = new AddressBook();
		addressBook1.name = "Kim Abercrombie";
		addressBook1.work = "1235550123";
		addressBook1.cell = "1235550123";
		addressBook1.home = "1235550123";
		addressBook1.email = "Kim@example.com";
		addressBook1.birthday = "4/13/1991";
		addressBook1.address = "123 N. Maple";
		addressBook1.city = "Cherryville";
		addressBook1.state = "WA";
		addressBook1.zip = "98031";
		addressBooks.add(addressBook1);

		AddressBook addressBook2 = new AddressBook();
		addressBook2.name = "John Smith";
		addressBook2.work = "3215230123";
		addressBook2.cell = "3215230123";
		addressBook2.home = "3215230123";
		addressBook2.email = "John@example.com";
		addressBook2.birthday = "5/20/1990";
		addressBook2.address = "4456 E. Aspen";
		addressBook2.city = "Montgomery";
		addressBook2.state = "AL";
		addressBook2.zip = "36136";
		addressBooks.add(addressBook2);

		AddressBook addressBook3 = new AddressBook();
		addressBook3.name = "James Williams";
		addressBook3.work = "5235550879";
		addressBook3.cell = "5235550879";
		addressBook3.home = "5235550879";
		addressBook3.email = "James@example.com";
		addressBook3.birthday = "4/5/1995";
		addressBook3.address = "123 N. Maple";
		addressBook3.city = "Denver";
		addressBook3.state = "CO";
		addressBook3.zip = "80214";
		addressBooks.add(addressBook3);

		AddressBook addressBook4 = new AddressBook();
		addressBook4.name = "Mark Jordan";
		addressBook4.work = "1238640185";
		addressBook4.cell = "1238640185";
		addressBook4.home = "1238640185";
		addressBook4.email = "Mark@example.com";
		addressBook4.birthday = "12/13/1988";
		addressBook4.address = "123 N. Maple";
		addressBook4.city = "Boise";
		addressBook4.state = "ID";
		addressBook4.zip = "83706";
		addressBooks.add(addressBook4);

		AddressBook addressBook5 = new AddressBook();
		addressBook5.name = "Andrew Lepp";
		addressBook5.work = "6235320178";
		addressBook5.cell = "6235320178";
		addressBook5.home = "6235320178";
		addressBook5.email = "Andrew@example.com";
		addressBook5.birthday = "10/9/1996";
		addressBook5.address = "123 N. Maple";
		addressBook5.city = "Augusta";
		addressBook5.state = "ME";
		addressBook5.zip = "04336";
		addressBooks.add(addressBook5);
		// #endregion

		// Add data source
		workbook.addDataSource("ds", addressBooks);

		// Invoke to process the template
		workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
		return "Template_TextFields.xlsx";
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
		return new String[] { "xlsx/Template_TextFields.xlsx" };
	}

	@Override
	public String[] getRefs() {
		return new String[] { "AddressBook" };
	}

    @Override
    public boolean getViewInGcPdfViewer() {
        return true;
    }

	@Override
	public boolean getCanDownload() {
		return false;
	}
}
