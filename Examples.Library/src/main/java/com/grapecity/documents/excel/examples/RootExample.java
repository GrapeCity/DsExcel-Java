package com.grapecity.documents.excel.examples;

public class RootExample extends FolderExample {
    public RootExample(String ns){
        super(ns);
    }

    @Override
    public String getNameResKey() {
        return "rootexample.name";
    }

    @Override
    public String getDescripResKey() {
        return "rootexample.descrip";
    }
}
