package com.grapecity.documents.excel.examples.templates.templatesamples;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FinancialDashboard extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_FinancialDashboard.xlsx");
        workbook.open(templateFile);

        ///#region Define custom classes
        //        public class FinancialRecord {
        //        	public String season;
        //        	public String country;
        //        	public double expect;
        //        	public double actual;
        //        }
        ///#endregion

        ///#region Init Data
        List<FinancialRecord> financialRecords = new ArrayList<FinancialRecord>();

        FinancialRecord financialRecord1 = new FinancialRecord();
        financialRecord1.season = "2016 Q1";
        financialRecord1.country = "USA";
        financialRecord1.expect = 236047;
        financialRecord1.actual = 328554;
        financialRecords.add(financialRecord1);

        FinancialRecord financialRecord2 = new FinancialRecord();
        financialRecord2.season = "2016 Q2";
        financialRecord2.country = "USA";
        financialRecord2.expect = 373060;
        financialRecord2.actual = 238136;
        financialRecords.add(financialRecord2);

        FinancialRecord financialRecord3 = new FinancialRecord();
        financialRecord3.season = "2016 Q3";
        financialRecord3.country = "USA";
        financialRecord3.expect = 224132;
        financialRecord3.actual = 300822;
        financialRecords.add(financialRecord3);

        FinancialRecord financialRecord4 = new FinancialRecord();
        financialRecord4.season = "2016 Q4";
        financialRecord4.country = "USA";
        financialRecord4.expect = 269305;
        financialRecord4.actual = 315337;
        financialRecords.add(financialRecord4);

        FinancialRecord financialRecord5 = new FinancialRecord();
        financialRecord5.season = "2017 Q1";
        financialRecord5.country = "USA";
        financialRecord5.expect = 265397;
        financialRecord5.actual = 279008;
        financialRecords.add(financialRecord5);

        FinancialRecord financialRecord6 = new FinancialRecord();
        financialRecord6.season = "2017 Q2";
        financialRecord6.country = "USA";
        financialRecord6.expect = 214079;
        financialRecord6.actual = 206019;
        financialRecords.add(financialRecord6);

        FinancialRecord financialRecord7 = new FinancialRecord();
        financialRecord7.season = "2017 Q3";
        financialRecord7.country = "USA";
        financialRecord7.expect = 370191;
        financialRecord7.actual = 238294;
        financialRecords.add(financialRecord7);

        FinancialRecord financialRecord8 = new FinancialRecord();
        financialRecord8.season = "2017 Q4";
        financialRecord8.country = "USA";
        financialRecord8.expect = 266843;
        financialRecord8.actual = 242323;
        financialRecords.add(financialRecord8);

        FinancialRecord financialRecord9 = new FinancialRecord();
        financialRecord9.season = "2016 Q1";
        financialRecord9.country = "Japan";
        financialRecord9.expect = 350156;
        financialRecord9.actual = 370834;
        financialRecords.add(financialRecord9);

        FinancialRecord financialRecord10 = new FinancialRecord();
        financialRecord10.season = "2016 Q2";
        financialRecord10.country = "Japan";
        financialRecord10.expect = 369399;
        financialRecord10.actual = 247324;
        financialRecords.add(financialRecord10);

        FinancialRecord financialRecord11 = new FinancialRecord();
        financialRecord11.season = "2016 Q3";
        financialRecord11.country = "Japan";
        financialRecord11.expect = 278834;
        financialRecord11.actual = 237385;
        financialRecords.add(financialRecord11);

        FinancialRecord financialRecord12 = new FinancialRecord();
        financialRecord12.season = "2016 Q4";
        financialRecord12.country = "Japan";
        financialRecord12.expect = 264277;
        financialRecord12.actual = 245048;
        financialRecords.add(financialRecord12);

        FinancialRecord financialRecord13 = new FinancialRecord();
        financialRecord13.season = "2017 Q1";
        financialRecord13.country = "Japan";
        financialRecord13.expect = 203006;
        financialRecord13.actual = 295389;
        financialRecords.add(financialRecord13);

        FinancialRecord financialRecord14 = new FinancialRecord();
        financialRecord14.season = "2017 Q2";
        financialRecord14.country = "Japan";
        financialRecord14.expect = 276987;
        financialRecord14.actual = 215804;
        financialRecords.add(financialRecord14);

        FinancialRecord financialRecord15 = new FinancialRecord();
        financialRecord15.season = "2017 Q3";
        financialRecord15.country = "Japan";
        financialRecord15.expect = 330315;
        financialRecord15.actual = 330443;
        financialRecords.add(financialRecord15);

        FinancialRecord financialRecord16 = new FinancialRecord();
        financialRecord16.season = "2017 Q4";
        financialRecord16.country = "Japan";
        financialRecord16.expect = 307477;
        financialRecord16.actual = 262512;
        financialRecords.add(financialRecord16);

        FinancialRecord financialRecord17 = new FinancialRecord();
        financialRecord17.season = "2016 Q1";
        financialRecord17.country = "Korea";
        financialRecord17.expect = 229432;
        financialRecord17.actual = 330368;
        financialRecords.add(financialRecord17);

        FinancialRecord financialRecord18 = new FinancialRecord();
        financialRecord18.season = "2016 Q2";
        financialRecord18.country = "Korea";
        financialRecord18.expect = 321904;
        financialRecord18.actual = 279114;
        financialRecords.add(financialRecord18);

        FinancialRecord financialRecord19 = new FinancialRecord();
        financialRecord19.season = "2016 Q3";
        financialRecord19.country = "Korea";
        financialRecord19.expect = 230496;
        financialRecord19.actual = 219257;
        financialRecords.add(financialRecord19);

        FinancialRecord financialRecord20 = new FinancialRecord();
        financialRecord20.season = "2016 Q4";
        financialRecord20.country = "Korea";
        financialRecord20.expect = 254328;
        financialRecord20.actual = 361880;
        financialRecords.add(financialRecord20);

        FinancialRecord financialRecord21 = new FinancialRecord();
        financialRecord21.season = "2017 Q1";
        financialRecord21.country = "Korea";
        financialRecord21.expect = 272263;
        financialRecord21.actual = 355419;
        financialRecords.add(financialRecord21);

        FinancialRecord financialRecord22 = new FinancialRecord();
        financialRecord22.season = "2017 Q2";
        financialRecord22.country = "Korea";
        financialRecord22.expect = 214079;
        financialRecord22.actual = 231510;
        financialRecords.add(financialRecord22);

        FinancialRecord financialRecord23 = new FinancialRecord();
        financialRecord23.season = "2017 Q3";
        financialRecord23.country = "Korea";
        financialRecord23.expect = 238392;
        financialRecord23.actual = 237430;
        financialRecords.add(financialRecord23);

        FinancialRecord financialRecord24 = new FinancialRecord();
        financialRecord24.season = "2017 Q4";
        financialRecord24.country = "Korea";
        financialRecord24.expect = 294097;
        financialRecord24.actual = 257680;
        financialRecords.add(financialRecord24);

        FinancialRecord financialRecord25 = new FinancialRecord();
        financialRecord25.season = "2016 Q1";
        financialRecord25.country = "China";
        financialRecord25.expect = 238175;
        financialRecord25.actual = 266070;
        financialRecords.add(financialRecord25);

        FinancialRecord financialRecord26 = new FinancialRecord();
        financialRecord26.season = "2016 Q2";
        financialRecord26.country = "China";
        financialRecord26.expect = 202721;
        financialRecord26.actual = 353563;
        financialRecords.add(financialRecord26);

        FinancialRecord financialRecord27 = new FinancialRecord();
        financialRecord27.season = "2016 Q3";
        financialRecord27.country = "China";
        financialRecord27.expect = 253279;
        financialRecord27.actual = 312586;
        financialRecords.add(financialRecord27);

        FinancialRecord financialRecord28 = new FinancialRecord();
        financialRecord28.season = "2016 Q4";
        financialRecord28.country = "China";
        financialRecord28.expect = 211847;
        financialRecord28.actual = 306970;
        financialRecords.add(financialRecord28);

        FinancialRecord financialRecord29 = new FinancialRecord();
        financialRecord29.season = "2017 Q1";
        financialRecord29.country = "China";
        financialRecord29.expect = 369314;
        financialRecord29.actual = 315718;
        financialRecords.add(financialRecord2);

        FinancialRecord financialRecord30 = new FinancialRecord();
        financialRecord30.season = "2017 Q2";
        financialRecord30.country = "China";
        financialRecord30.expect = 201224;
        financialRecord30.actual = 368630;
        financialRecords.add(financialRecord30);

        FinancialRecord financialRecord31 = new FinancialRecord();
        financialRecord31.season = "2017 Q3";
        financialRecord31.country = "China";
        financialRecord31.expect = 239792;
        financialRecord31.actual = 255108;
        financialRecords.add(financialRecord31);

        FinancialRecord financialRecord32 = new FinancialRecord();
        financialRecord32.season = "2017 Q4";
        financialRecord32.country = "China";
        financialRecord32.expect = 271096;
        financialRecord32.actual = 297354;
        financialRecords.add(financialRecord32);

        FinancialRecord financialRecord33 = new FinancialRecord();
        financialRecord33.season = "2016 Q1";
        financialRecord33.country = "India";
        financialRecord33.expect = 236047;
        financialRecord33.actual = 328554;
        financialRecords.add(financialRecord33);

        FinancialRecord financialRecord34 = new FinancialRecord();
        financialRecord34.season = "2016 Q2";
        financialRecord34.country = "India";
        financialRecord34.expect = 373060;
        financialRecord34.actual = 238136;
        financialRecords.add(financialRecord34);

        FinancialRecord financialRecord35 = new FinancialRecord();
        financialRecord35.season = "2016 Q3";
        financialRecord35.country = "India";
        financialRecord35.expect = 224132;
        financialRecord35.actual = 300822;
        financialRecords.add(financialRecord35);

        FinancialRecord financialRecord36 = new FinancialRecord();
        financialRecord36.season = "2016 Q4";
        financialRecord36.country = "India";
        financialRecord36.expect = 269305;
        financialRecord36.actual = 315337;
        financialRecords.add(financialRecord36);

        FinancialRecord financialRecord37 = new FinancialRecord();
        financialRecord37.season = "2017 Q1";
        financialRecord37.country = "India";
        financialRecord37.expect = 265397;
        financialRecord37.actual = 279008;
        financialRecords.add(financialRecord37);

        FinancialRecord financialRecord38 = new FinancialRecord();
        financialRecord38.season = "2017 Q2";
        financialRecord38.country = "India";
        financialRecord38.expect = 214079;
        financialRecord38.actual = 206019;
        financialRecords.add(financialRecord38);

        FinancialRecord financialRecord39 = new FinancialRecord();
        financialRecord39.season = "2017 Q3";
        financialRecord39.country = "India";
        financialRecord39.expect = 370191;
        financialRecord39.actual = 238294;
        financialRecords.add(financialRecord39);

        FinancialRecord financialRecord40 = new FinancialRecord();
        financialRecord40.season = "2017 Q4";
        financialRecord40.country = "India";
        financialRecord40.expect = 266843;
        financialRecord40.actual = 242323;
        financialRecords.add(financialRecord40);
        ///#endregion

        //Add data source
        workbook.addDataSource("ds", financialRecords);
        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_FinancialDashboard.xlsx";
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
        return new String[] { "xlsx/Template_FinancialDashboard.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
        return new String[] { "FinancialRecord" };
	}
}