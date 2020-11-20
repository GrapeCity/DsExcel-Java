package com.grapecity.documents.excel.examples;

import java.util.ArrayList;

public class Examples {

    private static FolderExample _rootExample;

    static
    {
        _rootExample = new RootExample(Examples.class.getPackage().getName());
        _rootExample.Level = 0;
    }

    public static FolderExample getRootExample()
    {
        return _rootExample;
    }

    public static ArrayList<ExampleBase> getAllExamples()
    {
        ArrayList<ExampleBase> examples = new ArrayList<ExampleBase>();
        examples.add(_rootExample);
        for (ExampleBase child : _rootExample.getChildren())
        {
            GetExamples(child, examples);
        }

        return examples;
    }

    private static void GetExamples(ExampleBase example, ArrayList<ExampleBase> examples)
    {
        examples.add(example);
        if (example instanceof FolderExample)
        {
            FolderExample folderExample = (FolderExample)example ;
            for (ExampleBase child : folderExample.getChildren())
            {
                GetExamples(child, examples);
            }
        }
    }
}
