package com.grapecity.documents.excel.examples;


import java.io.ByteArrayOutputStream;
import java.io.InputStream;

import javax.xml.bind.DatatypeConverter;

import com.grapecity.documents.excel.Workbook;

public class ExampleBase {


    public ExampleBase() {
    }

    public String getID() {
        return this.getClass().getName();
    }

    public String getNameByCulture(String culture) {
        return ResourceUtil.getResourceByCulture(culture, this.getNameResKey());
    }

    public String getDescriptionByCulture(String culture) {
        return ResourceUtil.getResourceByCulture(culture, this.getDescripResKey());
    }

    public String getCode() {
        return this.getExampleCode();
    }

    public String getCodeKotlin() {
        return this.getExampleCode_kotlin();
    }

    public boolean getCanDownload() {
        return true;
    }

    public boolean getShowViewer() {
        return true;
    }

    public boolean getShowScreenshot() {
        return false;
    }

    public boolean getShowCode() {
        return true;
    }

    public boolean getSavePdf() {
        return false;
    }
    
    public boolean getSavePageInfos() {
        return false;
    }
    
    public boolean getSaveAsImage() {
        return false;
    }

    public boolean getSaveCSV() {
        return false;
    }

    public boolean getHasTemplate() {
        return false;
    }

    public String getTemplateName() {
        return null;
    }

    public boolean getIsViewReadOnly() {
        return true;
    }

    public boolean getIsUpdate() {
        return false;
    }

    public boolean getIsNew() {
        return false;
    }

    public String getNameResKey() {
        return (this.getShortID() + ".Name").toLowerCase();
    }

    public String getDescripResKey() {
        return (this.getShortID() + ".Descrip").toLowerCase();
    }

    public String getShortID() {
        String id = this.getID();
        return id.substring(id.lastIndexOf(".") + 1);
    }

    public void executeExample(Workbook workbook, String userAgents) {
        this.beforeExecute(workbook, userAgents);
        this.execute(workbook);
        this.afterExecute(workbook, userAgents);
    }
    
    public void executeExample(ByteArrayOutputStream outputStream, Workbook workbook, String userAgents) {
        this.beforeExecute(workbook, userAgents);
        this.execute(workbook, outputStream);
        this.afterExecute(workbook, userAgents);
    }

    public void beforeExecute(Workbook workbook, String userAgents) {

    }

    public void execute(Workbook workbook) {

    }
    
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {

    }

    public void afterExecute(Workbook workbook, String userAgents) {
        if (this.getAgentIsMac(userAgents)) {
            //TODO, it will throw exception for now
            //workbook.calculate(); // ensure that all cached values can be saved in excel file, so number can display the file correctly even if the formulas are not supported in number.
        }
    }

    public String getScreenshotBase64() {
        String id = this.getID();
        String resourceName = id.toLowerCase().replace("com.grapecity.documents.excel.examples.features.", "");
        InputStream inputStream = this.getResourceStream(String.format("Screenshots/%s.png", resourceName));
        return "data:image/png;base64," + this.inputStreamToBase64String(inputStream);
    }

    private String getExampleCode() {
        String code = ResourceUtil.getCodeResource(this.getID());
        if (code != null && !code.isEmpty()) {
//            code = code.replaceAll("[\r\n][^\r\n]\\s{8}", "\n");
//            code = code.replaceAll("\\s{4}(.*)", "$1");
            code = this.processCodeFormat(code);
        }

        String codePre = "";
        if (this.getSavePageInfos()) {
        	codePre =  "  //create to a pdf file stream" + System.lineSeparator() + "  FileOutputStream outputStream = null;" +
        			   System.lineSeparator() + "  try {" + System.lineSeparator() + String.format("      outputStream = new FileOutputStream(\"%s.pdf\");", this.getShortID()) + 
        			   System.lineSeparator() + "  } catch (FileNotFoundException e) {" + 
        			   System.lineSeparator() + "      e.printStackTrace();" +
        			   System.lineSeparator() + "  }" + System.lineSeparator();
		}
        else if (this.getSaveAsImage()) {
        	codePre = "  //create to a png file stream" + System.lineSeparator() + "  FileOutputStream outputStream = null;" +
     			   System.lineSeparator() + "  try {" + System.lineSeparator() + String.format("      outputStream = new FileOutputStream(\"%s.png\");", this.getShortID()) + 
     			   System.lineSeparator() + "  } catch (FileNotFoundException e) {" + 
     			   System.lineSeparator() + "      e.printStackTrace();" +
     			   System.lineSeparator() + "  }" + System.lineSeparator();
		}
        
        codePre += "  //create a new workbook" + System.lineSeparator() + "  Workbook workbook = new Workbook();"; 
        
        String codePost = "";
        if (this.getSavePageInfos() || this.getSaveAsImage()) {
        	codePost = "  //close the file stream" + System.lineSeparator() + "  try {" + System.lineSeparator() +
        			   "      outputStream.close();" + System.lineSeparator() +
        			   "  } catch (IOException e) {" + System.lineSeparator() + 
        			   "      e.printStackTrace();" + System.lineSeparator() +
        			   "  }";
		} else if (this.getSavePdf()) {
            codePost = "  //save to an pdf file" + System.lineSeparator() + String.format("  workbook.save(\"%s.pdf\");", this.getShortID());
        } else if (this.getSaveCSV()) {
            codePost = "  //save to an csv file" + System.lineSeparator() + String.format("  workbook.save(\"%s.csv\");", this.getShortID());
        } else if (this.getCanDownload()) {
            codePost = "  //save to an excel file" + System.lineSeparator() + String.format("  workbook.save(\"%s.xlsx\");", this.getShortID());
        }
        code = codePre + code +  codePost;

        return code;
    }

    private String getExampleCode_kotlin() {
        String code = ResourceUtil.getCodeResource_kotlin(this.getID());
        if (code != null && !code.isEmpty()) {
            code = this.processCodeFormat(code);
        }

        String codePre = "";
        if (this.getSavePageInfos()) {
            codePre =  "  //create to a pdf file stream" + System.lineSeparator() + "  var outputStream: FileOutputStream? = null" +
                System.lineSeparator() + "  try {" + System.lineSeparator() + String.format("      outputStream = FileOutputStream(\"%s.pdf\")", this.getShortID()) +
                System.lineSeparator() + "  } catch (e: FileNotFoundException) {" +
                System.lineSeparator() + "      e.printStackTrace()" +
                System.lineSeparator() + "  }" + System.lineSeparator();
        }
        else if (this.getSaveAsImage()) {
            codePre = "  //create to a png file stream" + System.lineSeparator() + "  var outputStream: FileOutputStream? = null" +
                System.lineSeparator() + "  try {" + System.lineSeparator() + String.format("      outputStream = FileOutputStream(\"%s.png\")", this.getShortID()) +
                System.lineSeparator() + "  } catch (e: FileNotFoundException) {" +
                System.lineSeparator() + "      e.printStackTrace()" +
                System.lineSeparator() + "  }" + System.lineSeparator();
        }

        codePre += "  //create a new workbook" + System.lineSeparator() + "  var workbook = Workbook()";

        String codePost = "";
        if (this.getSavePageInfos() || this.getSaveAsImage()) {
            codePost = "  //close the file stream" + System.lineSeparator() + "  try {" + System.lineSeparator() +
                "       if(outputStream != null){"+ System.lineSeparator() +
                "           outputStream.close()" + System.lineSeparator() +
                "       }" + System.lineSeparator() + System.lineSeparator() +
                "  } catch (e: IOException) {" + System.lineSeparator() +
                "      e.printStackTrace()" + System.lineSeparator() +
                "  }";
        } else if (this.getSavePdf()) {
            codePost = "  //save to an pdf file" + System.lineSeparator() + String.format("  workbook.save(\"%s.pdf\")", this.getShortID());
        } else if (this.getSaveCSV()) {
            codePost = "  //save to an csv file" + System.lineSeparator() + String.format("  workbook.save(\"%s.csv\")", this.getShortID());
        } else if (this.getCanDownload()) {
            codePost = "  //save to an excel file" + System.lineSeparator() + String.format("  workbook.save(\"%s.xlsx\")", this.getShortID());
        }

        if(code == null){
            code = "";
        }
        code = codePre + code +  codePost;

        return code;
    }

    private String processCodeFormat(String code){
        String[] lines = code.split("\r\n|\n");
        StringBuilder builder = new StringBuilder();
        for(String line : lines){
            String s = line.replaceAll("^ {6}()", "$1");
            builder.append(s);
            builder.append(System.lineSeparator());
        }
        return builder.toString();
    }

    public boolean getAgentIsMac(String userAgents) {
        if (userAgents.toLowerCase().contains("macintosh")) {
            return true;
        }
        return false;

    }

    public InputStream getTemplateStream() {
        return this.getResourceStream("xlsx/" + this.getTemplateName());
    }

    public InputStream getResourceStream(String resource) {
        return ResourceUtil.getResourceStream(resource);
    }

    public String[] getResources(){
        return null;
    }

    public String[] getRefs(){
        return null;
    }

    private String inputStreamToBase64String(InputStream inputStream) {
        try {

            ByteArrayOutputStream byteArrayOutputStream = ResourceUtil.inputStreamToByteArrayOutputStream(inputStream);
            if (byteArrayOutputStream != null) {
                byte[] bytes = byteArrayOutputStream.toByteArray();
                String result = DatatypeConverter.printBase64Binary(bytes);
                return result;
            }
        } catch (Exception e) {
        }
        return null;
    }
}



