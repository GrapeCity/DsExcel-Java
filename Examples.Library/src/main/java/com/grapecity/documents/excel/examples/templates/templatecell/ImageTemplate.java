package com.grapecity.documents.excel.examples.templates.templatecell;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.IOUtils;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImageTemplate extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_ImageTemplate.xlsx");
        workbook.open(templateFile);

        //#region Define custom classes
        //public class BikeInfo
        //{
        //    public String bikeType;
        //    public List<BikeSeries> bikeSeries;
        //}

        //public class BikeSeries
        //{
        //    public String name;
        //    public String description;
        //    public byte[] bikeImage;
        //    public List<Bike> items;
        //}

        //public class Bike
        //{
        //    public String productNo;
        //    public String productName;
        //    public String color;
        //    public int size;
        //    public double weight;
        //    public double dealer;
        //    public double listPrice;
        //}
        //#endregion

        ///#region Init Data
        byte[] image1 = null;
        byte[] image2 = null;
        byte[] image3 = null;
        byte[] image4 = null;
        byte[] image5 = null;
        byte[] image6 = null;
        byte[] image7 = null;
        byte[] image8 = null;
        byte[] image9 = null;
        try {
        	InputStream image1Stream = this.getResourceStream("image/Mountain-100.jpg");
        	image1 = IOUtils.toByteArray(image1Stream);
        	InputStream image2Stream = this.getResourceStream("image/Mountain-200.jpg");
        	image2 = IOUtils.toByteArray(image2Stream);
        	InputStream image3Stream = this.getResourceStream("image/Mountain-300.jpg");
        	image3 = IOUtils.toByteArray(image3Stream);
        	InputStream image4Stream = this.getResourceStream("image/Mountain-400-W.jpg");
        	image4 = IOUtils.toByteArray(image4Stream);
        	InputStream image5Stream = this.getResourceStream("image/Mountain-500.jpg");
        	image5 = IOUtils.toByteArray(image5Stream);
        	InputStream image6Stream = this.getResourceStream("image/Road-150.jpg");
        	image6 = IOUtils.toByteArray(image6Stream);
        	InputStream image7Stream = this.getResourceStream("image/Road-350-W.jpg");
        	image7 = IOUtils.toByteArray(image7Stream);
        	InputStream image8Stream = this.getResourceStream("image/Touring-1000.jpg");
        	image8 = IOUtils.toByteArray(image8Stream);
        	InputStream image9Stream = this.getResourceStream("image/Touring-2000.jpg");
        	image9 = IOUtils.toByteArray(image9Stream);
        }
        catch (IOException e) {
        	e.printStackTrace();
        }

        List<BikeInfo> datasource = new ArrayList<BikeInfo>();

        BikeInfo bike1 = new BikeInfo();
        datasource.add(bike1);
        bike1.bikeType = "Mountain Bikes";
        bike1.bikeSeries = new ArrayList<BikeSeries>();

        BikeSeries bs1 = new BikeSeries();
        bike1.bikeSeries.add(bs1);
        bs1.name = "Mountain-100";
        bs1.bikeImage = image1;
        bs1.description = "Top-of-the-line competition mountain bike. Performance-enhancing options include the innovative HL Frame, super-smooth front suspension, and traction for all terrain.";
        bs1.items = new ArrayList<Bike>();
        Bike bItem1 = new Bike();
        bItem1.productNo = "BK-M82S-38";
        bItem1.productName = "Mountain-100 Silver, 38";
        bItem1.color = "Silver";
        bItem1.size = 38;
        bItem1.weight = 20.35;
        bItem1.dealer = 1912.1544;
        bItem1.listPrice = 3399.99;
        bs1.items.add(bItem1);
        Bike bItem2 = new Bike();
        bItem2.productNo = "BK-M82B-38";
        bItem2.productName = "Mountain-100 Black, 38";
        bItem2.color = "Black";
        bItem2.size = 38;
        bItem2.weight = 20.35;
        bItem2.dealer = 1898.0944;
        bItem2.listPrice = 3374.99;
        bs1.items.add(bItem2);

        BikeSeries bs2 = new BikeSeries();
        bike1.bikeSeries.add(bs2);
        bs2.name = "Mountain-200";
        bs2.bikeImage = image2;
        bs2.description = "Serious back-country riding. Perfect for all levels of competition. Uses the same HL Frame as the Mountain-100.";
        bs2.items = new ArrayList<Bike>();
        Bike bItem3 = new Bike();
        bItem3.productNo = "BK-M68S-42";
        bItem3.productName = "Mountain-200 Silver, 42";
        bItem3.color = "Silver";
        bItem3.size = 42;
        bItem3.weight = 23.77;
        bItem3.dealer = 1265.6195;
        bItem3.listPrice = 2319.99;
        bs2.items.add(bItem3);
        Bike bItem4 = new Bike();
        bItem4.productNo = "BK-M68B-38";
        bItem4.productName = "Mountain-200 Black, 38";
        bItem4.color = "Black";
        bItem4.size = 38;
        bItem4.weight = 23.35;
        bItem4.dealer = 1251.9813;
        bItem4.listPrice = 2294.99;
        bs2.items.add(bItem4);

        BikeSeries bs3 = new BikeSeries();
        bike1.bikeSeries.add(bs3);
        bs3.name = "Mountain-300";
        bs3.bikeImage = image3;
        bs3.description = "For true trail addicts.  An extremely durable bike that will go anywhere and keep you in control on challenging terrain - without breaking your budget.";
        bs3.items = new ArrayList<Bike>();
        Bike bItem5 = new Bike();
        bItem5.productNo = "BK-M47B-38";
        bItem5.productName = "Mountain-300 Black, 38";
        bItem5.color = "Black";
        bItem5.size = 38;
        bItem5.weight = 25.35;
        bItem5.dealer = 598.4354;
        bItem5.listPrice = 1079.99;
        bs3.items.add(bItem5);
        Bike bItem6 = new Bike();
        bItem6.productNo = "BK-M47B-40";
        bItem6.productName = "Mountain-300 Black, 40";
        bItem6.color = "Black";
        bItem6.size = 40;
        bItem6.weight = 25.77;
        bItem6.dealer = 598.4354;
        bItem6.listPrice = 1079.99;
        bs3.items.add(bItem6);

        BikeSeries bs4 = new BikeSeries();
        bike1.bikeSeries.add(bs4);
        bs4.name = "Mountain-400-W";
        bs4.bikeImage = image4;
        bs4.description = "This bike delivers a high-level of performance on a budget. It is responsive and maneuverable, and offers peace-of-mind when you decide to go off-road.";
        bs4.items = new ArrayList<Bike>();
        Bike bItem7 = new Bike();
        bItem7.productNo = "BKBK-M38S-38";
        bItem7.productName = "Mountain-400-W Silver, 38";
        bItem7.color = "Silver";
        bItem7.size = 38;
        bItem7.weight = 26.35;
        bItem7.dealer = 419.7784;
        bItem7.listPrice = 769.49;
        bs4.items.add(bItem7);
        Bike bItem8 = new Bike();
        bItem8.productNo = "BK-M38S-40";
        bItem8.productName = "Mountain-400-W Silver, 40";
        bItem8.color = "Silver";
        bItem8.size = 40;
        bItem8.weight = 26.77;
        bItem8.dealer = 419.7784;
        bItem8.listPrice = 769.49;
        bs4.items.add(bItem8);

        BikeSeries bs5 = new BikeSeries();
        bike1.bikeSeries.add(bs5);
        bs5.name = "Mountain-500";
        bs5.bikeImage = image5;
        bs5.description = "Suitable for any type of riding, on or off-road. Fits any budget. Smooth-shifting with a comfortable ride.";
        bs5.items = new ArrayList<Bike>();
        Bike bItem9 = new Bike();
        bItem9.productNo = "BK-M18S-40";
        bItem9.productName = "Mountain-500 Silver, 40";
        bItem9.color = "Silver";
        bItem9.size = 40;
        bItem9.weight = 27.35;
        bItem9.dealer = 308.2179;
        bItem9.listPrice = 564.99;
        bs5.items.add(bItem9);
        Bike bItem10 = new Bike();
        bItem10.productNo = "BK-M18B-40";
        bItem10.productName = "Mountain-500 Black, 40";
        bItem10.color = "Black";
        bItem10.size = 40;
        bItem10.weight = 27.35;
        bItem10.dealer = 294.5797;
        bItem10.listPrice = 539.99;
        bs5.items.add(bItem10);

        BikeInfo bike2 = new BikeInfo();
        datasource.add(bike2);
        bike2.bikeType = "Road Bikes";
        bike2.bikeSeries = new ArrayList<BikeSeries>();

        BikeSeries bs6 = new BikeSeries();
        bike2.bikeSeries.add(bs6);
        bs6.name = "Road-150";
        bs6.bikeImage = image6;
        bs6.description = "This bike is ridden by race winners. Developed with the Adventure Works Cycles professional race team, it has a extremely light heat-treated aluminum frame, and steering that allows precision control.";
        bs6.items = new ArrayList<Bike>();
        Bike bItem11 = new Bike();
        bItem11.productNo = "BK-R93R-62";
        bItem11.productName = "Road-150 Red, 62";
        bItem11.color = "Red";
        bItem11.size = 62;
        bItem11.weight = 15;
        bItem11.dealer = 2171.2942;
        bItem11.listPrice = 3578.27;
        bs6.items.add(bItem11);
        Bike bItem12 = new Bike();
        bItem12.productNo = "BK-R93R-44";
        bItem12.productName = "Road-150 Red, 44";
        bItem12.color = "Red";
        bItem12.size = 44;
        bItem12.weight = 13.77;
        bItem12.dealer = 2171.2942;
        bItem12.listPrice = 3578.27;
        bs6.items.add(bItem12);

        BikeSeries bs7 = new BikeSeries();
        bike2.bikeSeries.add(bs7);
        bs7.name = "Road-350-W";
        bs7.bikeImage = image7;
        bs7.description = "Cross-train, race, or just socialize on a sleek, aerodynamic bike designed for a woman.  Advanced seat technology provides comfort all day.";
        bs7.items = new ArrayList<Bike>();
        Bike bItem13 = new Bike();
        bItem13.productNo = "BK-R79Y-40";
        bItem13.productName = "Road-350-W Yellow, 40";
        bItem13.color = "Yellow";
        bItem13.size = 40;
        bItem13.weight = 15.35;
        bItem13.dealer = 1082.51;
        bItem13.listPrice = 1700.99;
        bs7.items.add(bItem13);
        Bike bItem14 = new Bike();
        bItem14.productNo = "BK-R79Y-42";
        bItem14.productName = "Road-350-W Yellow, 42";
        bItem14.color = "Yellow";
        bItem14.size = 42;
        bItem14.weight = 15.77;
        bItem14.dealer = 1082.51;
        bItem14.listPrice = 1700.99;
        bs7.items.add(bItem14);

        BikeInfo bike3 = new BikeInfo();
        datasource.add(bike3);
        bike3.bikeType = "Touring Bikes";
        bike3.bikeSeries = new ArrayList<BikeSeries>();

        BikeSeries bs8 = new BikeSeries();
        bike3.bikeSeries.add(bs8);
        bs8.name = "Touring-1000";
        bs8.bikeImage = image8;
        bs8.description = "Travel in style and comfort. Designed for maximum comfort and safety. Wide gear range takes on all hills. High-tech aluminum alloy construction provides durability without added weight.";
        bs8.items = new ArrayList<Bike>();
        Bike bItem15 = new Bike();
        bItem15.productNo = "BK-T79Y-46";
        bItem15.productName = "Touring-1000 Yellow, 46";
        bItem15.color = "Yellow";
        bItem15.size = 46;
        bItem15.weight = 25.13;
        bItem15.dealer = 1481.9379;
        bItem15.listPrice = 2384.07;
        bs8.items.add(bItem15);
        Bike bItem16 = new Bike();
        bItem16.productNo = "BK-T79U-46";
        bItem16.productName = "Touring-1000 Blue, 46";
        bItem16.color = "Blue";
        bItem16.size = 46;
        bItem16.weight = 25.13;
        bItem16.dealer = 1481.9379;
        bItem16.listPrice = 2384.07;
        bs8.items.add(bItem16);

        BikeSeries bs9 = new BikeSeries();
        bike3.bikeSeries.add(bs9);
        bs9.name = "Touring-2000";
        bs9.bikeImage = image9;
        bs9.description = "The plush custom saddle keeps you riding all day,  and there's plenty of space to add panniers and bike bags to the newly-redesigned carrier.  This bike has stability when fully-loaded.";
        bs9.items = new ArrayList<Bike>();
        Bike bItem17 = new Bike();
        bItem17.productNo = "BK-T44U-60";
        bItem17.productName = "Touring-2000 Blue, 60";
        bItem17.color = "Blue";
        bItem17.size = 60;
        bItem17.weight = 27.9;
        bItem17.dealer = 755.1508;
        bItem17.listPrice = 1214.85;
        bs9.items.add(bItem17);
        ///#endregion

        //Init template global settings
        workbook.getNames().add("TemplateOptions.KeepLineSize", "true");

        //Add data source
        workbook.addDataSource("ds", datasource);
        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_ImageTemplate.xlsx";
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
        return new String[] { "xlsx/Template_ImageTemplate.xlsx", "image/Mountain-100.jpg", "image/Mountain-200.jpg", "image/Mountain-300.jpg", "image/Mountain-400-W.jpg", "image/Mountain-500.jpg", "image/Road-150.jpg", "image/Road-350-W.jpg", "image/Touring-1000.jpg", "image/Touring-2000.jpg" };
	}

	@Override
	public String[] getRefs() {
        return new String[] { "BikeInfo", "BikeSeries", "Bike" };
	}

	@Override
	public boolean getIsNew() {
        return true;
	}
}