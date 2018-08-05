package com.grapecity.documents.excel.examples;

public class Examples {

    private static FolderExample _rootExample;

    static
    {
        _rootExample = new RootExample(Examples.class.getPackage().getName());
    }

    public static FolderExample getRootExample()
    {
        return _rootExample;
    }
}
