package com.grapecity.documents.excel.examples;

import java.lang.reflect.Array;
import java.lang.reflect.Constructor;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;

public class FolderExample extends ExampleBase {
    private ArrayList<ExampleBase> _chilren = null;
    private String _namespace;

    public FolderExample(String namespace) {
        _namespace = namespace;
    }

    @Override
    public String getID() {
        return _namespace;
    }

    @Override
    public String getNameResKey() {
        String shortName = _namespace.substring(_namespace.lastIndexOf(".") + 1, _namespace.length());
        return shortName + ".name";
    }

    @Override
    public String getDescripResKey() {
        String shortName = _namespace.substring(_namespace.lastIndexOf(".") + 1, _namespace.length());
        return shortName + ".descrip";
    }

    public ArrayList<ExampleBase> getChildren() {
        ArrayList<ExampleBase> children = new ArrayList<ExampleBase>();

        String[] types = ResourceUtil.listAllExamples();
        ArrayList<Class> classes = new ArrayList<Class>();

        try {
            for (String type : types) {
                Class exampleClass = Class.forName(type);
                if (exampleClass.getPackage().getName().startsWith(_namespace)) {
                    classes.add(exampleClass);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        HashSet<String> subNS = new HashSet<String>();

        for (Class exampelClass : classes) {
            String packageName = exampelClass.getPackage().getName();
            if (packageName.equals(_namespace)) {
                Constructor<ExampleBase> childConstructor = exampelClass.getConstructors()[0];
                try {

                    ExampleBase child = childConstructor.newInstance();
                    children.add(child);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            } else if (!subNS.contains(packageName)) {
                String ends = packageName.substring(this._namespace.length() + 1);
                if (ends != null && !ends.isEmpty()) {
                    String[] nsItems = ends.split("\\.");
                    String currentNS = _namespace + "." + nsItems[0];
                    if (!subNS.contains(currentNS)) {
                        children.add(new FolderExample(currentNS));
                        subNS.add(currentNS);
                    }
                    subNS.add(packageName);
                }
            }
        }

        Collections.sort(children, new ExampleComparer());

        return children;
    }

    public ExampleBase findExample(String id) {
        return this.findExample(this, id);
    }

    private ExampleBase findExample(ExampleBase example, String id) {
        if (example.getID().equals(id)) {
            return example;
        }

        if(example instanceof FolderExample) {
            FolderExample folderExample = (FolderExample) example;
            if (folderExample != null) {
                for (ExampleBase child : folderExample.getChildren()) {
                    ExampleBase result = this.findExample(child, id);
                    if (result != null) {
                        return result;
                    }
                }
            }
        }

        return null;
    }

    @Override
    public boolean getIsNew() {
        return false;
    }

    @Override
    public boolean getIsUpdate() {
        return this.isUpdateRecursive(this);
    }

    private boolean isUpdateRecursive(ExampleBase example) {
        if (example instanceof FolderExample) {
            FolderExample childFolderExample = (FolderExample) example;
            for (ExampleBase item : childFolderExample.getChildren()) {
                if (item.getIsUpdate() || item.getIsNew()) {
                    return true;
                }

                if (isUpdateRecursive(item)) {
                    return true;
                }
            }
        } else if (example.getIsUpdate() || example.getIsNew()) {
            return true;
        }

        return false;
    }
}
